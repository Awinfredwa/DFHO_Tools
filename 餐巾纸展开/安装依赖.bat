@echo off
chcp 65001 >nul
echo.
echo ========================================
echo     🔧 餐巾纸展开处理器 - 依赖安装
echo ========================================
echo.
echo 正在安装所需的 Python 库...
echo.

echo 📦 安装 Pillow (图片处理库)...
pip install Pillow

if errorlevel 1 (
    echo.
    echo ❌ 安装失败！
    echo.
    echo 可能的解决方案：
    echo 1. 确保已安装 Python 3.6 或更新版本
    echo 2. 确保 pip 可用
    echo 3. 检查网络连接
    echo.
    echo 手动安装命令：
    echo pip install Pillow
    echo.
    pause
    exit /b 1
)

echo.
echo ✅ 所有依赖安装完成！
echo.
echo 现在您可以运行"餐巾纸展开.bat"开始处理图片了。
echo.
pause 