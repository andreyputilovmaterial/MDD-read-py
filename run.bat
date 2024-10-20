@ECHO OFF


SET "MDD=examples\R2301349M.mdd"


ECHO -
ECHO 1. read
ECHO read from: %MDD%
ECHO write to: .json
python read_mdd.py --mdd "%MDD%"
if %ERRORLEVEL% NEQ 0 ( echo ERROR: Failure && pause && exit /b %errorlevel% )

ECHO -
ECHO 2. generate html
set "read_mdd_json=%MDD%.json"
python lib\MDM-Report-py\report_create.py "%read_mdd_json%"
if %ERRORLEVEL% NEQ 0 ( echo ERROR: Failure && pause && exit /b %errorlevel% )

ECHO done!

