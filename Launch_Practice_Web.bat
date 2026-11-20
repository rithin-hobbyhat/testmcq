@echo off
title HOBBY HAT MCQ - Web & Mobile Server
cd /d "%~dp0"

echo ========================================================
echo        Starting HOBBY HAT MCQ Server...
echo ========================================================
echo.

py server.py
if %errorlevel% neq 0 (
    python server.py
)

pause
