@echo off
chcp 65001 >nul
title PDF 工具箱
python pdf_toolbox_gui.py
if errorlevel 1 (
    echo.
    echo 程序运行失败，请先双击 install.bat 安装依赖。
    echo.
    pause
)
