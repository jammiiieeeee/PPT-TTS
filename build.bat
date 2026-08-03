@echo off
cd /d "%~dp0"
title Building PowerPoint Narrator

REM Setup proper ANSI escape character for colors
for /F %%a in ('echo prompt $E ^| cmd') do set "ESC=%%a"
set "GREEN=%ESC%[92m"
set "RED=%ESC%[91m"
set "YELLOW=%ESC%[93m"
set "CYAN=%ESC%[96m"
set "RESET=%ESC%[0m"

echo.
echo %CYAN%========================================%RESET%
echo %CYAN%   Building PowerPoint Narrator%RESET%
echo %CYAN%========================================%RESET%
echo.

REM ── Check for spec file ───────────────────────────────────────────────────
if not exist PPT-TTS.spec (
    echo %RED%ERROR: PPT-TTS.spec not found.%RESET%
    echo Run this script from the project root directory.
    pause
    exit /b 1
)

REM ── Install build dependencies ────────────────────────────────────────────
echo %YELLOW%Installing build dependencies...%RESET%
python -m pip install pyinstaller python-pptx boto3 lxml pywin32 --quiet --user 2>nul
if %errorlevel% neq 0 (
    echo %RED%ERROR: Failed to install build dependencies.%RESET%
    pause
    exit /b 1
)
echo %GREEN%Dependencies ready.%RESET%

REM ── Build ─────────────────────────────────────────────────────────────────
echo.
echo %YELLOW%Building executable...%RESET%
python -m PyInstaller PPT-TTS.spec --clean --noconfirm
if %errorlevel% neq 0 (
    echo.
    echo %RED%ERROR: Build failed. Check the output above for errors.%RESET%
    pause
    exit /b 1
)

echo.
echo %GREEN%Build complete!%RESET%
echo.
echo Output: dist\PPT-TTS.exe
echo.
echo To distribute:
echo   1. Copy dist\PPT-TTS.exe
echo   2. Copy config.json (with your AWS credentials)
echo   3. Place both in the same folder on the target machine
echo.
pause
