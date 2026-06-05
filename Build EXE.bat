@echo off
REM ============================================================
REM  Builds "Electricity Bill Merger.exe" — a standalone double-
REM  clickable application with custom icon. Run this ONCE.
REM
REM  After it finishes, the .exe will be in the "dist" folder
REM  next to this file. You can copy that .exe anywhere and
REM  double-click to run — no Python needed on the target PC.
REM ============================================================
cd /d "%~dp0"

echo.
echo === Electricity Bill Merger — EXE Builder ===
echo.

REM 1. Locate Python
where python >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python is not installed or not on PATH.
    echo Please install Python 3.8+ from https://www.python.org first.
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

REM 2. Install / upgrade PyInstaller + required runtime libs
echo Installing PyInstaller, pandas, openpyxl ...
python -m pip install --upgrade pip -q
python -m pip install --upgrade pyinstaller pandas openpyxl -q
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to install required packages.
    pause
    exit /b 1
)

REM 3. Build the .exe
echo.
echo Building Electricity Bill Merger.exe ...
echo (This can take 1-2 minutes the first time.)
echo.

python -m PyInstaller ^
    --noconfirm ^
    --onefile ^
    --windowed ^
    --name "Electricity Bill Merger" ^
    --icon "app_icon.ico" ^
    --hidden-import openpyxl ^
    --hidden-import pandas ^
    "merge_app.py"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Build failed. Scroll up to see the error message.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  BUILD SUCCESSFUL
echo  Your application is at:
echo     %CD%\dist\Electricity Bill Merger.exe
echo.
echo  You can copy that single .exe to your Desktop and
echo  double-click it. No Python needed on the target machine.
echo ============================================================
echo.
pause
