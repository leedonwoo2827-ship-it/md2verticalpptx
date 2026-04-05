@echo off
chcp 65001 >nul

if not exist ".venv\Scripts\activate.bat" (
    echo [ERROR] Virtual environment not found. Run install.bat first.
    pause
    exit /b 1
)

call .venv\Scripts\activate.bat
python -m md2pptx %*
