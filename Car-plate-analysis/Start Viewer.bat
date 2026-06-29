@echo off
title Car Plate Viewer
echo Stopping any existing server on port 5000...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":5000 " ^| findstr "LISTENING"') do (
    taskkill /PID %%a /F >nul 2>&1
)
timeout /t 1 /nobreak > nul
echo Starting Car Plate Viewer...
start /b py "%~dp0app.py"
timeout /t 2 /nobreak > nul
echo Opening browser...
start "" http://127.0.0.1:5000
echo.
echo Server is running. Close this window to stop.
pause > nul
