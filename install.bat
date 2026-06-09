@echo off
title PDF Toolbox - Install Dependencies

echo ========================================
echo   PDF Toolbox - Install Dependencies
echo ========================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python was not found.
    echo Please install Python 3.8 or later, and enable Add Python to PATH.
    echo Download: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo [1/3] Checking Python version...
python --version

echo.
echo [2/3] Upgrading pip...
python -m pip install --upgrade pip
if errorlevel 1 (
    echo [WARN] Failed to upgrade pip. Continue installing dependencies...
)

echo.
echo [3/3] Installing dependencies from requirements.txt...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [ERROR] Failed to install dependencies.
    echo Please check your network connection and Python environment.
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Installation completed!
echo ========================================
echo.
echo Start GUI: double-click run_gui.bat
echo Or run: python pdf_toolbox_gui.py
echo.
echo Note: Office to PDF requires LibreOffice.
echo Download: https://www.libreoffice.org/download/download-libreoffice/
echo.
pause
