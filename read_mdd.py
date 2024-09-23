# import os, time, re, sys
from datetime import datetime, timezone
# from dateutil import tz
import argparse
from pathlib import Path
import re
import json


# import pythoncom
import win32com.client


# TODO:
import pdb



class MDMDocument:
    
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
        config_default = {
            'features': ['label','properties','translations'], # ,'scripting'],
            'contexts': ['Question','Analysis']
        }
        self.__config = { **config_default, **config }
    
    def __del__(self):
        self.__document.Close()
        print('MDM document closed')
    
    def read(self):
        self.__translations = [ '{langcode}'.format(langcode=langcode) for langcode in self.__document.Languages ]
        result = {
            'MDD': self.__mdd_path,
            'run_time_utc': self.__read_datetime.replace(tzinfo=timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
            'run_time_local': self.__read_datetime.strftime('%Y-%m-%dT%H:%M:%SZ'),
            'properties': self.__read_properties(),
            'languages': self.__read_languages(),
            'types': self.__read_sharedlists(),
            'fields': self.__read_fields(self.__document.Fields),
            'pages': self.__read_pages(),
            'routing': self.__read_routing()
        }
        return result
    
    def __read_properties(self):
        
        try:

            result = []

            config = self.__config
            document = self.__document

            for read_feature in config['features']:
                if read_feature=='label':
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
                for read_feature in config['features']:
                    if read_feature=='label':
                        result.append(item.LongName)

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
                        'elements': [],
                    },
                    **self.__read_mdm_item(item)
                }
                for cat in item.Elements:
                    cat_name = '{name}'.format(name=cat.Name)
                    result_item['elements'].append({
                    **{
                        'name': cat_name
                    },
                    **self.__read_mdm_item(cat)
                })
                result.append(
                    result_item
                )

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
                        'name': item_name,
                        'fields': [],
                    },
                    **self.__read_mdm_item(item)
                }
                for cat in item:
                    cat_name = '{name}'.format(name=cat.Name)
                    result_item['fields'].append({
                    **{
                        'name': cat_name
                    },
                    **self.__read_mdm_item(cat)
                })
                result.append(
                    result_item
                )

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
                    result.append(
                        result_item
                    )

                except Exception as e:
                    print('failed when processing "{name}"'.format(name=item_name))
                    raise e
        
            return result
        
        except Exception as e:
            print('failed when processing fields')
            raise e
    
    def __read_process_field(self,item):

        item_name = item.Name
        try:

            result_item = {
                **{
                    'name': '{name}'.format(name=item.Name),
                    'object_type_value': item.ObjectTypeValue,
                    #'data_type': item.DataType,
                    #'is_grid': item.IsGrid,
                },
                **self.__read_mdm_item(item)
            }
            object_type_value = item.ObjectTypeValue
            if object_type_value==0:
                # regular variable
                result_item['type'] = 'plain'
                data_type = item.DataType
                result_item['data_type'] = data_type
                if data_type==0:
                    # info
                    result_item['type'] = 'plain/info'
                elif data_type==1:
                    # long
                    result_item['type'] = 'plain/long'
                    result_item['minvalue'] = item.MinValue
                    result_item['maxvalue'] = item.MaxValue
                elif data_type==2:
                    # text
                    result_item['type'] = 'plain/text'
                    result_item['minvalue'] = item.MinValue
                    result_item['maxvalue'] = item.MaxValue
                elif data_type==3:
                    # categorical
                    result_item['type'] = 'plain/categorical'
                    result_item['minvalue'] = item.MinValue
                    result_item['maxvalue'] = item.MaxValue
                    result_item['categories'] = []
                    for cat in item.Categories:
                        result_item['categories'].append(self.__read_mdm_item(cat))
                elif data_type==5:
                    # date
                    result_item['type'] = 'plain/date'
                elif data_type==6:
                    # double
                    result_item['type'] = 'plain/double'
                    result_item['minvalue'] = item.MinValue
                    result_item['maxvalue'] = item.MaxValue
                elif data_type==7:
                    # boolean
                    result_item['type'] = 'plain/boolean'
                pass
            elif object_type_value==1:
                # array (loop)
                result_item['type'] = 'array'
                result_item['is_grid'] = item.IsGrid
                result_item['categories'] = []
                for cat in item.Categories:
                    result_item['categories'].append(self.__read_mdm_item(cat))
                result_item['fields'] = []
                for cat in item.Fields:
                    result_item['fields'].append(self.__read_process_field(cat))
            elif object_type_value==2:
                # Grid (it seems it's something different than Array, but I can't understand their logic; maybe it's different because it has a different db setup in case data, I don't know)
                # Execute Error: The '<Object>.IGrid' type does not support the 'categories' property
                result_item['type'] = 'grid'
                result_item['is_grid'] = item.IsGrid
                result_item['categories'] = []
                for cat in item.Elements:
                    result_item['categories'].append(self.__read_mdm_item(cat))
                result_item['fields'] = []
                for cat in item.Fields:
                    result_item['fields'].append(self.__read_process_field(cat))
            elif object_type_value==3:
                # class (block)
                result_item['type'] = 'block'
                result_item['fields'] = []
                for cat in item.Fields:
                    result_item['fields'].append(self.__read_process_field(cat))
            elif object_type_value==16:
                result_item['type'] = 'plain16'
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
            return result_item
        
        except Exception as e:
            print('failed when processing "{name}"'.format(name=item_name))
            raise e
    
    def __read_routing(self):

        try:

            result = None

            config = self.__config
            document = self.__document

            for routing_part in ['Web']: # ??? TODO:
                result = '{val}'.format(val=document.Routing.Script)

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
                elif read_feature=='translations':
                #elif read_feature[:9]=='langcode-':
                    #langcode = read_feature[9:]
                    for langcode in self.__translations:
                        #val_label = '{val}'.format(val=item.Labels["Label"].Text["Question"][langcode])
                        # TODO:
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
        metavar='p123456.mdd',
        help='Input MDD',
        required=True
    )
    parser.add_argument(
        '-2',
        '--method',
        metavar='open',
        help='Method',
        required=False
    )
    args = parser.parse_args()
    inp_mdd = None
    if args.mdd:
        inp_mdd = Path(args.mdd)
        inp_mdd = '{inp_mdd}'.format(inp_mdd=inp_mdd.resolve())
    # inp_file_specs = open(inp_file_specs_name, encoding="utf8")

    method = '{arg}'.format(arg=args.method) if args.method else 'open'

    print('MDM read script: script started at {dt}'.format(dt=time_start))

    MDMDocument = MDMDocument(inp_mdd,method)

    result = MDMDocument.read()
    
    result_json = json.dumps(result, indent=4)
    result_json_fname = re.sub(r'^\s*?(.*?)\s*?$',lambda m: '{base}{added}'.format(base=m[1],added='.json'),'{path}'.format(path=inp_mdd))
    print('MDM read script: saving as "{fname}"'.format(fname=result_json_fname))
    with open(result_json_fname, "w") as outfile:
        outfile.write(result_json)

    del MDMDocument

    time_finish = datetime.now()
    print('MDM read script: finished at {dt} (elapsed {duration})'.format(dt=time_finish,duration=time_finish-time_start))

