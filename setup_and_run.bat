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

echo.
echo Installing dependencies...
%PYTHON% -m pip install -r requirements.txt --quiet
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

REM Validate config.json has real credentials
%PYTHON% -c "import json,sys; c=json.load(open('config.json')); aws=c.get('aws',{}); k=aws.get('access_key_id',''); s=aws.get('secret_access_key',''); sys.exit(0 if k and s and not k.startswith('YOUR_') and not s.startswith('YOUR_') else 1)"
if %errorlevel% neq 0 (
    echo.
    echo ERROR: config.json is missing valid AWS credentials.
    echo.
    echo Opening config.json for you to edit - replace the placeholder values.
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
