@echo off
setlocal
cd /d "%~dp0"
title Napkin Image Tool

echo.
echo    Napkin Image Tool
echo.

for %%F in (*.py) do (
    python "%%F"
    if errorlevel 1 goto run_failed
    goto done
)

echo No Python script found in this folder.
pause
exit /b 1

:done
pause
exit /b 0

:run_failed
echo.
echo Tool failed. Please run the setup batch file first.
echo If it still fails, make sure Python and Pillow are installed.
echo.
pause
exit /b 1
