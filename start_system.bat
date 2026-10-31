@echo off
title Exam Student Matcher & Discrepancy Finder
cd /d "%~dp0"

echo ===================================================
echo   Starting Exam Student Matcher Web Server...
echo ===================================================
echo.

python run_server.py

pause
