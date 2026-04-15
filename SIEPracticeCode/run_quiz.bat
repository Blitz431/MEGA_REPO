@echo off
cd /d "%~dp0"

:: Check Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Python is not installed. Download it from https://python.org
    pause
    exit /b 1
)

:: Auto-install openpyxl if missing (silent, fast if already installed)
python -m pip install openpyxl --quiet --disable-pip-version-check

python quiz.py
pause
