@echo off
cd /d %~dp0
echo =========================================
echo Installing BurnIn Monitor dependencies
echo =========================================
echo.

"C:\Users\kevin_hsieh\AppData\Local\Python\bin\python3.14.exe" -m pip install -r requirements.txt
if errorlevel 1 (
    echo [Error] Dependency installation failed.
    echo Please check network connection and pip configuration.
    pause
    exit /b 1
)

echo.
echo Done. You can run run.bat to start the program.
echo Tesseract OCR installer:
echo https://github.com/UB-Mannheim/tesseract/wiki
pause
