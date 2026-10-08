@echo off
REM ============================================================
REM  Corneal Reports — one-click Windows build
REM
REM  Produces:  dist\CornealReportsApp.exe  (single-file, no console)
REM  Requires:  Python 3.9+ on PATH (python / pip)
REM ============================================================

setlocal ENABLEDELAYEDEXPANSION
cd /d "%~dp0"

echo.
echo === Corneal Reports — building Windows executable ===
echo Working dir: %CD%
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not on PATH.
    echo         Install Python 3.9+ from https://www.python.org/downloads/
    echo         and tick "Add python.exe to PATH" during install.
    pause
    exit /b 1
)

echo [1/3] Installing/upgrading dependencies...
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)

echo.
echo [2/3] Cleaning previous build...
if exist build  rmdir /s /q build
if exist dist   rmdir /s /q dist

echo.
echo [3/3] Running PyInstaller...
python -m PyInstaller --noconfirm --clean corneal_reports_app.spec
if errorlevel 1 (
    echo [ERROR] PyInstaller failed.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  BUILD SUCCESSFUL
echo ============================================================
echo  Executable : %CD%\dist\CornealReportsApp.exe
echo.
echo  Usage:
echo    1. Copy dist\CornealReportsApp.exe anywhere you like.
echo    2. Put your PDF reports in a folder named "reports"
echo       next to the .exe (or pick a different folder inside the app).
echo    3. Double-click CornealReportsApp.exe.
echo ============================================================
echo.
pause
endlocal
