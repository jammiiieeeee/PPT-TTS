@echo off
echo ===================================
echo   Building PPT-TTS executable
echo ===================================

REM Install build dependencies (use python -m pip to avoid Device Guard blocks)
python -m pip install pyinstaller python-pptx boto3 lxml pywin32 --user

REM Build
python -m PyInstaller PPT-TTS.spec --clean --noconfirm

echo.
echo Build complete! Output: dist\PPT-TTS.exe
echo Copy both PPT-TTS.exe and config.json to distribute.
pause
