@echo off
chcp 65001 >nul 2>&1
echo ========================================
echo    XML Comparator - Launch
echo ========================================
echo.
echo Monitoring: Status updates every 5 min
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    pause
    exit /b 1
)

REM Install dependencies if needed
echo Installing dependencies...
python -m pip install --quiet --upgrade pip
python -m pip install --quiet -r requirements.txt

if %errorlevel% neq 0 (
    echo.
    echo ERROR during dependencies installation
    pause
    exit /b 1
)

echo.
echo Launching XML comparator...
echo Processing will display status every 5 minutes
echo.

REM Launch the Python script
python xml_comparator.py

REM Pause at the end to see results
pause
