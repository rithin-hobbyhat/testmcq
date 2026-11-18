@echo off
title MOQ Question Bank Authoring Studio
cd /d "%~dp0"
echo ========================================================
echo   MOQ Question Bank Authoring Studio
echo   Starting application...
echo ========================================================
py app.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Application exited with an error. Press any key to close this window...
    pause >nul
)
