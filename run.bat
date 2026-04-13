@echo off
cd /d %~dp0
"C:\Users\kevin_hsieh\AppData\Local\Python\bin\python3.14.exe" burnin_monitor.py
if errorlevel 1 (
    echo.
    echo [Error] Failed to start burnin_monitor.py
)
pause
