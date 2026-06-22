@echo off
chcp 65001 >nul
title 餐巾纸套图生成工具

echo.
echo ====================================================
echo    餐巾纸套图生成工具
echo ====================================================
echo.

python "生成餐巾纸套图.py"

if errorlevel 1 (
    echo.
    echo ❌ 错误：Python 可能未安装，或 Pillow 未安装。
    echo 请先双击 "配置环境.bat"。
    echo.
    pause
)
