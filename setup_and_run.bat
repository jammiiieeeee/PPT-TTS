@echo off
cd /d "%~dp0"
title PowerPoint Narrator

REM Setup proper ANSI escape character for colors
for /F %%a in ('echo prompt $E ^| cmd') do set "ESC=%%a"
set "GREEN=%ESC%[92m"
set "RED=%ESC%[91m"
set "YELLOW=%ESC%[93m"
set "CYAN=%ESC%[96m"
set "RESET=%ESC%[0m"

echo.
echo %CYAN%========================================%RESET%
echo %CYAN%   PowerPoint Narrator%RESET%
echo %CYAN%========================================%RESET%
echo.

REM ── Detect Python ─────────────────────────────────────────────────────────
where python >nul 2>&1
if %errorlevel% equ 0 (
    set "PYTHON=python"
    goto :check_deps
)
where py >nul 2>&1
if %errorlevel% equ 0 (
    set "PYTHON=py"
    goto :check_deps
)

echo %RED%ERROR: Python is not installed or not in PATH.%RESET%
echo.
echo Install Python from: https://www.python.org/downloads/
echo Check "Add Python to PATH" during install.
echo.
pause
exit /b 1

:check_deps
echo %YELLOW%Checking dependencies...%RESET%

REM Check if dependencies are already installed
%PYTHON% -c "import pptx, boto3, lxml, windnd" >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing dependencies (this may take a minute^)...
    %PYTHON% -m pip install -r requirements.txt --quiet --disable-pip-version-check 2>nul
    if %errorlevel% neq 0 (
        echo.
        echo %RED%ERROR: Failed to install dependencies.%RESET%
        echo Try running: pip install -r requirements.txt
        pause
        exit /b 1
    )
    echo %GREEN%Dependencies installed.%RESET%
) else (
    echo %GREEN%Dependencies ready.%RESET%
)

REM ── Check AWS credentials ─────────────────────────────────────────────────
if not exist config.json (
    echo.
    echo %YELLOW%Creating config.json template...%RESET%
    %PYTHON% -c "import json; json.dump({'aws': {'access_key_id': 'YOUR_AWS_ACCESS_KEY_ID', 'secret_access_key': 'YOUR_AWS_SECRET_ACCESS_KEY', 'region': 'us-east-1'}}, open('config.json', 'w'), indent=4)"
)

%PYTHON% -c "import json,sys; c=json.load(open('config.json')); aws=c.get('aws',{}); k=aws.get('access_key_id',''); s=aws.get('secret_access_key',''); sys.exit(0 if k and s and not k.startswith('YOUR_') and not s.startswith('YOUR_') else 1)" 2>nul
if %errorlevel% neq 0 (
    echo.
    echo %RED%ERROR: AWS credentials not configured.%RESET%
    echo.
    echo Open config.json and add your AWS credentials:
    echo   - Access Key ID
    echo   - Secret Access Key
    echo.
    echo Get these from: AWS Console ^> IAM ^> Users ^> Security credentials
    echo.
    notepad config.json
    echo.
    echo Save the file and run this script again.
    pause
    exit /b 1
)

echo %GREEN%AWS credentials configured.%RESET%

REM ── Test connection ───────────────────────────────────────────────────────
echo.
echo %YELLOW%Testing AWS connection...%RESET%
%PYTHON% -c "import json,boto3; c=json.load(open('config.json')); aws=c['aws']; kw={'aws_access_key_id':aws['access_key_id'],'aws_secret_access_key':aws['secret_access_key'],'region_name':aws.get('region','us-east-1')}; v=aws.get('verify_ssl',False); kw['verify']=v; sts=boto3.client('sts',**kw); r=sts.get_caller_identity(); print(f'Authenticated as: {r[\"Arn\"]}')" 2>nul
if %errorlevel% neq 0 (
    echo.
    echo %RED%ERROR: AWS credentials are invalid or expired.%RESET%
    echo.
    echo Please check your credentials in config.json.
    echo.
    notepad config.json
    pause
    exit /b 1
)

echo.
echo %GREEN%All checks passed!%RESET%

REM ── Create desktop shortcut (first run only) ──────────────────────────────
set "SCRIPT_DIR=%~dp0"
set "SHORTCUT_PATH=%USERPROFILE%\Desktop\PowerPoint Narrator.lnk"
if not exist "%SHORTCUT_PATH%" (
    echo.
    echo %YELLOW%Creating desktop shortcut...%RESET%
    powershell -NoProfile -Command "$s=(New-Object -COM WScript.Shell).CreateShortcut('%SHORTCUT_PATH%'); $s.TargetPath='%SCRIPT_DIR%generate_audio_ppt.py'; $s.WorkingDirectory='%SCRIPT_DIR%'; $s.Description='PowerPoint Narrator'; $s.Save()"
    if exist "%SHORTCUT_PATH%" (
        echo %GREEN%Desktop shortcut created.%RESET%
        echo.
        echo %CYAN%TIP: You can now run PowerPoint Narrator directly from your desktop.%RESET%
        echo %CYAN%     This setup script is only needed for the first run.%RESET%
    ) else (
        echo %YELLOW%Could not create shortcut. You can still run this script directly.%RESET%
    )
)

echo.
echo %CYAN%Starting PowerPoint Narrator...%RESET%
echo.
%PYTHON% generate_audio_ppt.py
