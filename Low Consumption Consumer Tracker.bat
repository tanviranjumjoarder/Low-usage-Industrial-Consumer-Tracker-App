@echo off
REM ============================================================
REM  Double-click launcher for Electricity Bill Merger
REM  (Shows a small console window in the background)
REM ============================================================
cd /d "%~dp0"

REM Try pythonw (no console) first, fall back to python
where pythonw >nul 2>nul
if %ERRORLEVEL%==0 (
    start "" pythonw "merge_app.py"
    exit /b
)

where python >nul 2>nul
if %ERRORLEVEL%==0 (
    python "merge_app.py"
    exit /b
)

echo.
echo Python was not found on this computer.
echo Please install Python 3.8+ from https://www.python.org
echo.
pause
