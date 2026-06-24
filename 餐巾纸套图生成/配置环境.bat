@echo off
setlocal
cd /d "%~dp0"
title Setup

echo.
echo Installing required Python packages...
echo.

python -m pip install -r requirements.txt
if errorlevel 1 goto install_failed

echo.
echo    Setup complete
echo.
echo You can now run the tool.
echo.
pause
exit /b 0

:install_failed
echo.
echo Setup failed. Please make sure Python is installed.
echo.
pause
exit /b 1
