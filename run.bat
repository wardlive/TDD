@echo off
setlocal DisableDelayedExpansion
chcp 65001 >nul
set "PYTHONUTF8=1"
set "TDD_PROJECT_DIR=%~dp0"
start "TDD" "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -Command "try { $ErrorActionPreference = 'Stop'; Set-Location -LiteralPath $env:TDD_PROJECT_DIR; if (-not (Test-Path -LiteralPath '.\.venv\Scripts\python.exe' -PathType Leaf)) { throw 'Сначала настройте программу по инструкции в README.md.' }; & '.\.venv\Scripts\python.exe' main.py; $code = $LASTEXITCODE; if ($code -ne 0) { Read-Host 'Нажмите Enter, чтобы закрыть окно' }; exit $code } catch { Write-Host $_ -ForegroundColor Red; Read-Host 'Нажмите Enter, чтобы закрыть окно'; exit 1 }"
exit /b %ERRORLEVEL%
