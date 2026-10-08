@echo off
REM Run the app directly with the system Python (no .exe build needed).
setlocal
cd /d "%~dp0"

where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not on PATH. Install Python 3.9+ from python.org.
    pause
    exit /b 1
)

python -m pip install -r requirements.txt >nul 2>&1
python corneal_reports_app.py
endlocal
