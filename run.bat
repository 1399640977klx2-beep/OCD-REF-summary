@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

set VENV_DIR=venv

echo ============================================
echo  Project Bootstrap
echo ============================================

:: --------------------  Check Python  --------------------
where python >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Python is not found in PATH.
    echo         Please make sure Python is installed and added to PATH.
    pause
    exit /b 1
)
python --version

:: --------------------  Virtual Environment  --------------------
if exist "%VENV_DIR%\Scripts\activate.bat" (
    echo [INFO] Virtual environment found, activating...
) else (
    echo [INFO] No virtual environment detected, creating ...
    python -m venv "%VENV_DIR%"
    if %ERRORLEVEL% neq 0 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
    echo [OK] Virtual environment created.
)

call "%VENV_DIR%\Scripts\activate.bat"

:: --------------------  Install Dependencies  --------------------
if exist requirements.txt (
    echo [INFO] Installing dependencies from requirements.txt ...
    pip install -r requirements.txt
    if %ERRORLEVEL% neq 0 (
        echo [WARN] pip install encountered errors, trying to continue...
    )
) else (
    echo [INFO] No requirements.txt found, skipping dependency install.
)

:: --------------------  Run  --------------------
echo [INFO] Starting main.py ...
python main.py

pause
