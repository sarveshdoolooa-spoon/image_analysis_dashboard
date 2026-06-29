@echo off
title Car Plate Viewer (Streamlit)
echo Stopping any existing server on ports 8501 / 8765...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8501 " ^| findstr "LISTENING"') do (
    taskkill /PID %%a /F >nul 2>&1
)
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8765 " ^| findstr "LISTENING"') do (
    taskkill /PID %%a /F >nul 2>&1
)
timeout /t 1 /nobreak > nul
echo Starting Car Plate Viewer...
cd /d "%~dp0"
start "" http://127.0.0.1:8501
py -m streamlit run "%~dp0streamlit_app.py" --server.headless true
pause
