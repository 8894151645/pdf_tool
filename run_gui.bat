@echo off
title PDF Toolbox
python pdf_toolbox_gui.py
if errorlevel 1 (
    echo.
    echo Failed to start the program.
    echo Please run install.bat first to install dependencies.
    echo.
    pause
)
