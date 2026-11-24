@echo off
title PDF Site Request Extractor - Daily Excel Appender
echo ========================================================
echo   Starting PDF Site Request Extractor UI...
echo   Please wait while the dashboard opens in your browser.
echo ========================================================
echo.
cd /d "%~dp0"
python -m streamlit run app.py --server.headless false
pause
