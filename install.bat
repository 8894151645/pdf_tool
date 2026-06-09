@echo off
chcp 65001 >nul
title PDF 工具箱 - 自动安装依赖

echo ========================================
echo   PDF 工具箱 - 自动安装依赖
echo ========================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [错误] 未检测到 Python。
    echo 请先安装 Python 3.9 或更高版本，并勾选 Add Python to PATH。
    echo 下载地址：https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo [1/3] 检查 Python 版本...
python --version

echo.
echo [2/3] 升级 pip...
python -m pip install --upgrade pip
if errorlevel 1 (
    echo [警告] pip 升级失败，继续尝试安装依赖...
)

echo.
echo [3/3] 安装 requirements.txt 中的依赖...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [错误] 依赖安装失败，请检查网络或 Python 环境。
    pause
    exit /b 1
)

echo.
echo ========================================
echo   安装完成！
echo ========================================
echo.
echo 运行图形界面：双击 run_gui.bat
echo 或者执行：python pdf_toolbox_gui.py
echo.
echo 注意：Word/PPT 转 PDF 需要额外安装 LibreOffice。
echo 下载地址：https://www.libreoffice.org/download/download-libreoffice/
echo.
pause
