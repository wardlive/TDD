@echo off
setlocal
chcp 65001 >nul
set "PYTHONUTF8=1"
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Сначала настройте программу по инструкции в README.md, раздел «Как запустить на Windows».
    pause
    exit /b 1
)
".venv\Scripts\python.exe" main.py
set "TDD_EXIT_CODE=%ERRORLEVEL%"
if not "%TDD_EXIT_CODE%"=="0" pause
exit /b %TDD_EXIT_CODE%
