@echo off
title MCQ Practice Arena - Web & Mobile Server
cd /d "%~dp0"

echo ========================================================
echo        Starting MCQ Practice Arena Server...
echo ========================================================
echo.

py server.py
if %errorlevel% neq 0 (
    python server.py
)

pause
