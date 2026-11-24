@echo off
title Frontier Towers Philippines - Automated SRF Extractor
echo ========================================================
echo   Frontier Tower Associates Philippines (FTAP)
echo   Starting Automated SRF Extractor Dashboard...
echo   Please wait while the dashboard opens in your browser.
echo ========================================================
echo.
cd /d "%~dp0"

IF EXIST ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -m streamlit run app.py --server.headless false
) ELSE IF EXIST "%USERPROFILE%\.local\bin\uv.exe" (
    "%USERPROFILE%\.local\bin\uv.exe" run streamlit run app.py --server.headless false
) ELSE (
    python -m streamlit run app.py --server.headless false
)
pause
