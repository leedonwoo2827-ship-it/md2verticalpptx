@echo off
chcp 65001 >nul
echo ============================================
echo   md2pptx Installer (Windows)
echo ============================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Install Python 3.10+ first.
    echo   https://www.python.org/downloads/
    pause
    exit /b 1
)

:: Check Python version >= 3.10
for /f "tokens=2 delims= " %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo [OK] Python %PYVER%

:: Create venv
if not exist ".venv" (
    echo Creating virtual environment...
    python -m venv .venv
)

:: Activate and install
echo Installing dependencies...
call .venv\Scripts\activate.bat
pip install --upgrade pip >nul 2>&1
pip install -r requirements.txt
pip install -e .

echo.
echo ============================================
echo   Installation complete!
echo ============================================
echo.
echo Usage:
echo   run.bat ^<body.md^> -t ^<templates_dir^> [-o output.pptx]
echo.
pause
