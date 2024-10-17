# import os, time, re, sys
from datetime import datetime, timezone
# from dateutil import tz
import argparse
from pathlib import Path
import re
import json


# import pythoncom
import win32com.client


# # TODO:
# import pdb



def normaize_linebreaks(s):
    return re.sub(r'(?:(?:\r)|(?:\n))+',"\n",s)




# how to use this class
# 1. create an object of type MDMDocument (and pass the path to MDD as a param)
# the MDD is opened as a file system object, so the connection is made to the file (descriptor is acquired)
# also, the third param is to pass config options - which properties to read (for example, just labes, or translations too), which sections, etc...
# I don't have cofig options documented, please read through source code to understand config power
# 2. call "read"
# so the program iterates over all fields
# and returns result as dict that can be represented as json
# that's it, enjoy

# when done, file descriptor is released (this is also happening on exceptions because it is specced in __del__ method)

class MDMDocument:
    

    # constructor: load MDD document here
    # three params:
    # 1. the path to MDD,
    # 2. and an optional method to read: "open" or "join";
    #    both are working options, just different methods used;
    #    I am not sure which method is better
    #    as we know, FileTrimmer scripts are using the "join" method
    #    and PrepData transfomrations too
    # 3. config options (I don't have cofig options documented, please read through source code to understand config power)
    def __init__(self,mdd_path,method='open',config={}):

        self.__document = None

        if method=='open':
            mDocument = win32com.client.Dispatch("MDM.Document")
            # openConstants_oNOSAVE = 3
            openConstants_oREAD = 1
            # openConstants_oREADWRITE = 2
            print('opening MDM document using method "open": "{path}"'.format(path=mdd_path))
            mDocument.Open( mdd_path, "", openConstants_oREAD )
            self.__document = mDocument
        elif method=='join':
            mDocument = win32com.client.Dispatch("MDM.Document")
            print('opening MDM document using method "join": "{path}"'.format(path=mdd_path))
            mDocument.Join(mdd_path, "{..}", 1, 32|16|512)
            self.__document = mDocument
        else:
            raise Exception('MDM Open: Unknown open method, {method}'.format(method=method))
        
        self.__mdd_path = mdd_path
        self.__read_datetime = datetime.now()
        # config is initialized here if it was not passed from above; the default config is passed and is set below! Do not update here! Like, you can update here, but anyway it is passed from below!
        config_default = {
            'features': ['label','attributes','properties','translations','scripting'],
            'contexts': ['Question','Analysis']
        }
        self.__config = { **config_default, **config }
    


    # unlink document if some error happened, or if we are done processing it
    def __del__(self):
        self.__document.Close()
        print('MDM document closed')



    # strange methods required by python so that I can use "with"
    # I still don't understand why this is needed as we already have __init__ and __del__ and allll should work, why on Earth __enter__ and __exit__ are necessary????
    def __enter__(self):
        return self    
    def __exit__(self, exc_type, exc_val, exc_tb):
        pass
    


    # actually the main method of this class
    # reads the entire MDD and return it as a "report"
    # return type is dict that is suitable to be saved or sent as json
    def read(self):

        # prep some variables - list of languages, list of features, columns, etc
        translations_list = [ '{langcode}'.format(langcode=langcode) for langcode in self.__document.Languages ]
        self.__translations = translations_list
        flags_list = []
        columns_list = ['name']
        for feature_spec in self.__config['features']:
            col_add = None
            if feature_spec=='label':
                col_add = 'label'
            elif feature_spec=='attributes':
                col_add = 'attributes'
            elif feature_spec=='properties':
                col_add = 'properties'
            elif feature_spec=='scripting':
                col_add = 'scripting'
            elif feature_spec=='translations':
                col_add = None # TODO: not super beautiful design sorry
                for langcode in translations_list:
                    columns_list.append('langcode-{langcode}'.format(langcode=langcode))
            else:
                raise Exception('feature type is not recognized: "{ft}"'.format(ft=feature_spec))
            if  col_add:
                columns_list.append(col_add)
        
        # ok, here's the final result
        # that's what we return
        result = {
            'report_type': 'MDD',
            'source_file': '{path}'.format(path=self.__mdd_path),
            'report_datetime_utc': self.__read_datetime.astimezone(tz=timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
            'report_datetime_local': self.__read_datetime.strftime('%Y-%m-%dT%H:%M:%SZ'),
            'source_file_metadata': [
                { 'name': 'MDD', 'value': '{path}'.format(path=self.__mdd_path) },
            ],
            'report_scheme': {
                'columns': columns_list,
                'flags': flags_list
            },
            'sections': [
                { 'name': 'mdmproperties', 'content': [{'name':'MDM','properties':self.__read_properties()}] },
                { 'name': 'languages', 'content': self.__read_languages() },
                { 'name': 'shared_lists', 'content': self.__read_sharedlists() },
                { 'name': 'fields', 'content': self.__read_fields(self.__document.Fields) },
                { 'name': 'pages', 'content': self.__read_pages() },
                { 'name': 'routing', 'content': self.__read_routing() },
            ],
        }
        return result
    
    # and individual methods to follow: to read through all properties, all shared lists, all fields, categories, pages, etc...
    def __read_properties(self):
        
        try:

            result = []

            config = self.__config
            document = self.__document

            for read_feature in config['features']:
                if read_feature=='label':
                    pass
                elif read_feature=='attributes':
                    pass
                elif read_feature=='properties':
                    item = document
                    context_preserve = document.Contexts.Current
                    properties_list = []
                    properties = {}
                    for read_context in document.Contexts:
                        if '{ctx}'.format(ctx=read_context).lower() in [ctx.lower() for ctx in config['contexts']]:
                            document.Contexts.Current = read_context
                            for index_prop in range( 0, item.Properties.Count ):
                                prop_name = '{name}'.format(name=item.Properties.Name(index_prop))
                                properties_list.append(prop_name)
                                properties[prop_name] = '{value}'.format(value=item.Properties[prop_name])
                    document.Contexts.Current = context_preserve
                    for prop_name in properties_list:
                        result.append({ 'name': prop_name, 'value': properties[prop_name] })
            return result
        
        except Exception as e:
            print('failed when processing properties')
            raise e
    
    def __read_languages(self):

        try:

            result = []

            config = self.__config
            document = self.__document

            for item in document.Languages:
                result_part = {'name':'{name}'.format(name=item.Name)}
                for read_feature in config['features']:
                    if read_feature=='label':
                        result_part['label'] = '{val}'.format(val=item.LongName)
                result.append(result_part)

            return result
        
        except Exception as e:
            print('failed when processing languages')
            raise e
    
    def __read_sharedlists(self):

        try:

            result = []

            # config = self.__config
            document = self.__document
            fields = document.types

            sharedlists_list = [ '{slname}'.format(slname=slname.Name) for slname in fields ]
            # TODO: sort
            for sl_name in sharedlists_list:
                item = fields[sl_name]
                result_item = {
                    **{
                        'name': sl_name,
                        # 'elements': [],
                    },
                    **self.__read_mdm_item(item)
                }
                result.append(
                    result_item
                )
                result_item = None
                for cat in item.Elements:
                    #cat_name = '{name}'.format(name=cat.Name)
                    element_item = self.__read_mdm_item(cat)
                    element_item['name'] = '{prefix}.categories[{itemname}]'.format(prefix=sl_name,itemname=element_item['name'])
                    result.append( element_item )

            return result
        
        except Exception as e:
            print('failed when processing shared lists')
            raise e
    
    def __read_pages(self):

        try:

            result = []

            # config = self.__config
            document = self.__document
            fields = document.pages

            pages_list = [ '{name}'.format(name=slname.Name) for slname in fields ]
            # TODO: sort
            for item_name in pages_list:
                item = fields[item_name]
                result_item = {
                    **{
                        'name': item_name
                    },
                    **self.__read_mdm_item(item)
                }
                result.append(
                    result_item
                )
                result_item = None
                for cat in item:
                    #cat_name = '{name}'.format(name=cat.Name)
                    item_add = self.__read_mdm_item(cat)
                    item_add['name'] = '{prefix}.{name}'.format(prefix=item_name,name=item_add['name'])
                    result.append(item_add)

            return result
        
        except Exception as e:
            print('failed when processing pages')
            raise e
    
    def __read_fields(self,fields):

        try:

            result = []

            # config = self.__config

            fields_list = [ '{name}'.format(name=item.Name) for item in fields ]
            # TODO: sort
            for item_name in fields_list:
                try:
                    item = fields[item_name]
                    result_item = self.__read_process_field(item)
                    #result.append( result_item )
                    result = result + result_item

                except Exception as e:
                    print('failed when processing "{name}"'.format(name=item_name))
                    raise e
        
            return result
        
        except Exception as e:
            print('failed when processing fields')
            raise e
    
    def __read_process_field(self,item):

        item_name = item.Name
        result_other_items = []
        try:

            result_item = {
                **{
                    'name': '{name}'.format(name=item_name),
                    'attributes': {
                        'object_type_value': item.ObjectTypeValue,
                        #'data_type': item.DataType,
                        #'is_grid': item.IsGrid,
                    },
                },
                **self.__read_mdm_item(item)
            }
            object_type_value = item.ObjectTypeValue
            if object_type_value==0:
                # regular variable
                result_item['attributes']['type'] = 'plain'
                data_type = item.DataType
                result_item['attributes']['data_type'] = data_type
                if data_type==0:
                    # info
                    result_item['attributes']['type'] = 'plain/info'
                elif data_type==1:
                    # long
                    result_item['attributes']['type'] = 'plain/long'
                    result_item['attributes']['minvalue'] = item.MinValue
                    result_item['attributes']['maxvalue'] = item.MaxValue
                elif data_type==2:
                    # text
                    result_item['attributes']['type'] = 'plain/text'
                    result_item['attributes']['minvalue'] = item.MinValue
                    result_item['attributes']['maxvalue'] = item.MaxValue
                elif data_type==3:
                    # categorical
                    result_item['attributes']['type'] = 'plain/categorical'
                    result_item['attributes']['minvalue'] = item.MinValue
                    result_item['attributes']['maxvalue'] = item.MaxValue
                    for cat in item.Categories:
                        item_add = self.__read_mdm_item(cat)
                        item_add['name'] = '{prefix}.categories[{name}]'.format(prefix=item_name,name=item_add['name'])
                        result_other_items.append(item_add)
                elif data_type==5:
                    # date
                    result_item['attributes']['type'] = 'plain/date'
                elif data_type==6:
                    # double
                    result_item['attributes']['type'] = 'plain/double'
                    result_item['attributes']['minvalue'] = item.MinValue
                    result_item['attributes']['maxvalue'] = item.MaxValue
                elif data_type==7:
                    # boolean
                    result_item['attributes']['type'] = 'plain/boolean'
                pass
            elif object_type_value==1:
                # array (loop)
                result_item['attributes']['type'] = 'array'
                result_item['attributes']['is_grid'] = item.IsGrid
                for cat in item.Categories:
                        item_add = self.__read_mdm_item(cat)
                        item_add['name'] = '{prefix}.categories[{name}]'.format(prefix=item_name,name=item_add['name'])
                        result_other_items.append(item_add)
                for cat in item.Fields:
                    #result_item['attributes']['fields'].append(self.__read_process_field(cat))
                    result_other_items = result_other_items + [ {**item,'name':'{prefix}.{part}'.format(prefix=item_name,part=item['name'])} for item in self.__read_process_field(cat) ]
            elif object_type_value==2:
                # Grid (it seems it's something different than Array, but I can't understand their logic; maybe it's different because it has a different db setup in case data, I don't know)
                # Execute Error: The '<Object>.IGrid' type does not support the 'categories' property
                result_item['attributes']['type'] = 'grid'
                result_item['attributes']['is_grid'] = item.IsGrid
                for cat in item.Elements:
                        item_add = self.__read_mdm_item(cat)
                        item_add['name'] = '{prefix}.categories[{name}]'.format(prefix=item_name,name=item_add['name'])
                        result_other_items.append(item_add)
                for cat in item.Fields:
                    #result_item['attributes']['fields'].append(self.__read_process_field(cat))
                    result_other_items = result_other_items + [ {**item,'name':'{prefix}.{part}'.format(prefix=item_name,part=item['name'])} for item in self.__read_process_field(cat) ]
            elif object_type_value==3:
                # class (block)
                result_item['attributes']['type'] = 'block'
                #result_item['attributes']['fields'] = []
                for cat in item.Fields:
                    #result_item['attributes']['fields'].append(self.__read_process_field(cat))
                    result_other_items = result_other_items + [ {**item,'name':'{prefix}.{part}'.format(prefix=item_name,part=item['name'])} for item in self.__read_process_field(cat) ]
            elif object_type_value==16:
                result_item['attributes']['type'] = 'plain/type16'
                # not sure what is it, an example is Respondent.Serial (in some projects)
                pass
            else:
                raise ValueError('unrecognized object data type: {val}'.format(val=object_type_value))
            # for cat in item:
            #     cat_name = '{name}'.format(name=cat.Name)
            #     result_item['fields'].append({
            #     **{
            #         'name': cat_name
            #     },
            #     **self.__read_mdm_item(cat)
            # })

            # we need to reformat attributes collection
            attributes_upd = []
            for itemKey in result_item['attributes'].keys():
                attributes_upd.append({'name':itemKey,'value':'{val}'.format(val=result_item['attributes'][itemKey])})
            result_item['attributes'] = attributes_upd

            return [result_item] + result_other_items
        
        except Exception as e:
            print('failed when processing "{name}"'.format(name=item_name))
            raise e
    
    def __read_routing(self):

        try:

            result = []

            config = self.__config
            document = self.__document

            for routing_part in ['DefaultRoutingContext']:
                val = '{val}'.format(val=document.Routing.Script)
                result.append({'name':routing_part,'label':val})

            return result
        
        except Exception as e:
            print('failed when processing routing')
            raise e
        
    def __read_mdm_item(self,item):

        item_name = '{name}'.format(name=item.Name)

        try:

            result = {
                'name': item_name
            }

            config = self.__config
            document = self.__document

            for read_feature in config['features']:
                if read_feature=='label':
                    val_label = '{val}'.format(val=item.Label)
                    result[read_feature] = val_label
                elif read_feature=='attributes':
                    pass
                elif read_feature=='translations':
                #elif read_feature[:9]=='langcode-':
                    #langcode = read_feature[9:]
                    for langcode in self.__translations:
                        #val_label = '{val}'.format(val=item.Labels["Label"].Text["Question"][langcode])
                        val_label = '{val}'.format(val='???')
                        # item.Labels('Question','ENU')
                        try:
                            val_label = '{val}'.format(val=item.Labels('Question',langcode))
                        except Exception as e:
                            val_label = '{val}'.format(val=e)
                        result['langcode-{langcode}'.format(langcode=langcode)] = val_label
                elif read_feature=='properties':
                    result_properties = []
                    context_preserve = document.Contexts.Current
                    properties_list = []
                    properties = {}
                    for read_context in document.Contexts:
                        if '{ctx}'.format(ctx=read_context).lower() in [ctx.lower() for ctx in config['contexts']]:
                            document.Contexts.Current = read_context
                            for index_prop in range( 0, item.Properties.Count ):
                                prop_name = '{name}'.format(name=item.Properties.Name(index_prop))
                                properties_list.append(prop_name)
                                properties[prop_name] = '{value}'.format(value=item.Properties[prop_name])
                    document.Contexts.Current = context_preserve
                    for prop_name in properties_list:
                        result_properties.append({ 'name': prop_name, 'value': properties[prop_name] })
                    result[read_feature] = result_properties
                elif read_feature=='scripting':
                    #val_label = '{val}'.format(val=item.Script)
                    # TODO:
                    val_label = '{val}'.format(val='???')
                    try:
                        val_label = '{val}'.format(val=item.Script)
                    except Exception as e:
                        val_label = '{val}'.format(val=e)
                    result[read_feature] = val_label
                else:
                    raise ValueError('options param is not supported: {feature_type}'.format(feature_type=read_feature))
            return result
        
        except Exception as e:
            print('failed when processing "{name}"'.format(name=item_name))
            raise e



if __name__ == '__main__':
    time_start = datetime.now()
    parser = argparse.ArgumentParser(
        description="Read MDD"
    )
    parser.add_argument(
        '-1',
        '--mdd',
        help='Input MDD',
        required=True
    )
    parser.add_argument(
        '--method',
        default='open',
        help='Method',
        required=False
    )
    parser.add_argument(
        '--config-features',
        help='Config: list features (default is label,properties,translations)',
        required=False
    )
    parser.add_argument(
        '--config-contexts',
        help='Config: list contexts (default is Question,Analysis)',
        required=False
    )
    args = parser.parse_args()
    inp_mdd = None
    if args.mdd:
        inp_mdd = Path(args.mdd)
        inp_mdd = '{inp_mdd}'.format(inp_mdd=inp_mdd.resolve())
    # inp_file_specs = open(inp_file_specs_name, encoding="utf8")

    method = '{arg}'.format(arg=args.method) if args.method else 'open'

    config = {
        # 'features': ['label','attributes','properties','translations'], # ,'scripting'],
        'features': ['label','attributes','properties','translations','scripting'],
        'contexts': ['Question','Analysis']
    }
    if args.config_features:
        config['features'] = args.config_features.split(',')
    if args.config_contexts:
        config['contexts'] = args.config_contexts.split(',')

    print('MDM read script: script started at {dt}'.format(dt=time_start))

    with MDMDocument(inp_mdd,method,config) as doc:

        result = doc.read()
        
        result_json = json.dumps(result, indent=4)
        result_json_fname = ( Path(inp_mdd).parents[0] / '{basename}{ext}'.format(basename=Path(inp_mdd).name,ext='.json') if Path(inp_mdd).is_file() else re.sub(r'^\s*?(.*?)\s*?$',lambda m: '{base}{added}'.format(base=m[1],added='.json'),'{path}'.format(path=inp_mdd)) )
        print('MDM read script: saving as "{fname}"'.format(fname=result_json_fname))
        with open(result_json_fname, "w") as outfile:
            outfile.write(result_json)

    time_finish = datetime.now()
    print('MDM read script: finished at {dt} (elapsed {duration})'.format(dt=time_finish,duration=time_finish-time_start))

