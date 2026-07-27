@echo off
cd /d "%~dp0"
echo ===================================
echo   PPT TTS - First Time Setup
echo ===================================

REM Detect Python command
where python >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON=python
    goto :found
)
where py >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON=py
    goto :found
)

echo.
echo ERROR: Python is not installed or not in PATH.
echo.
echo Install Python from https://www.python.org/downloads/
echo IMPORTANT: Check "Add Python to PATH" during install, then run this script again.
pause
exit /b 1

:found
echo Found: %PYTHON%
%PYTHON% --version

%PYTHON% -m pip install -r requirements.txt --quiet --disable-pip-version-check 2>nul
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Failed to install dependencies.
    pause
    exit /b 1
)

REM Create config.json from template if it doesn't exist
if not exist config.json (
    %PYTHON% -c "import json; json.dump({'aws': {'access_key_id': 'YOUR_AWS_ACCESS_KEY_ID', 'secret_access_key': 'YOUR_AWS_SECRET_ACCESS_KEY', 'region': 'us-east-1'}}, open('config.json', 'w'), indent=4)"
)

REM Validate config.json and test AWS credentials
%PYTHON% -c "import json,sys; c=json.load(open('config.json')); aws=c.get('aws',{}); k=aws.get('access_key_id',''); s=aws.get('secret_access_key',''); sys.exit(0 if k and s and not k.startswith('YOUR_') and not s.startswith('YOUR_') else 1)" 2>nul
if %errorlevel% neq 0 (
    echo.
    echo ERROR: config.json is missing AWS credentials.
    echo.
    echo Please make sure config.json has your AWS Access Key ID and Secret Access Key.
    echo You can get these from the AWS Console under IAM > Users > Security credentials.
    echo.
    echo Opening config.json for you to edit...
    notepad config.json
    echo.
    echo After saving the file, run this script again.
    pause
    exit /b 1
)

%PYTHON% -c "import json,boto3; c=json.load(open('config.json')); aws=c['aws']; kw={'aws_access_key_id':aws['access_key_id'],'aws_secret_access_key':aws['secret_access_key'],'region_name':aws.get('region','us-east-1')}; v=aws.get('verify_ssl',False); kw['verify']=v; sts=boto3.client('sts',**kw); r=sts.get_caller_identity(); print(f'Authenticated as: {r[\"Arn\"]}')" 2>nul
if %errorlevel% neq 0 (
    echo.
    echo ERROR: AWS credentials are invalid or expired.
    echo.
    echo Please check that your Access Key ID and Secret Access Key are correct.
    echo Keys can be regenerated in the AWS Console under IAM > Users > Security credentials.
    echo.
    echo Opening config.json for you to edit...
    notepad config.json
    echo.
    echo After saving the file, run this script again.
    pause
    exit /b 1
)

echo.
echo Starting PPT TTS...
echo.
%PYTHON% generate_audio_ppt.py

pause
