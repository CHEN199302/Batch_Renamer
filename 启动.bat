@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo    Batch File Rename Tool (with Path)
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python not found. Please install Python 3.6+
    echo.
    pause
    exit /b
)

REM Check and install dependencies
echo Checking dependencies...
python -c "import openpyxl" 2>nul
if errorlevel 1 (
    echo Installing openpyxl...
    pip install openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple
)

echo.
echo Starting rename tool...
echo.

REM Use default config
python batch_rename.py

REM If you need to use custom config, uncomment the line below
REM python batch_rename.py --config config_custom.ini

echo.
pause