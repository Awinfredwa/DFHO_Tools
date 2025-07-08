@echo off
chcp 65001 >nul
echo.
echo ========================================
echo     🧾 餐巾纸展开处理器
echo ========================================
echo.
echo 正在启动餐巾纸展开脚本...
echo.

python "餐巾纸展开.py"

if errorlevel 1 (
    echo.
    echo ❌ 错误：Python 可能未安装或脚本执行失败
    echo 请确保已安装 Python 和所需的库 (PIL/Pillow)
    echo.
    echo 安装命令：
    echo pip install Pillow
    echo.
    echo 或运行："安装依赖.bat"
    echo.
    pause
) 