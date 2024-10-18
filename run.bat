@ECHO OFF


SET "MDD=examples\R2301349M.mdd"


ECHO -
ECHO 1. read
ECHO read from: %MDD%
ECHO write to: .json
python MDD.py --mdd "%MDD%"
if %ERRORLEVEL% NEQ 0 ( echo ERROR: Failure && pause && exit /b %errorlevel% )

ECHO -
ECHO 2. generate html
set "read_mdd_json=%MDD%.json"
python lib\MDM-HTMLReport-py\report_create.py "%read_mdd_json%"
if %ERRORLEVEL% NEQ 0 ( echo ERROR: Failure && pause && exit /b %errorlevel% )

ECHO -
ECHO 3. diff!
ECHO ... TBD
if %ERRORLEVEL% NEQ 0 ( echo ERROR: Failure && pause && exit /b %errorlevel% )

ECHO done!

