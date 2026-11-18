@echo off
setlocal DisableDelayedExpansion
rem Keep this file ASCII-only: cmd misparses Cyrillic in .bat files (chcp 65001 does not help).
set "PYTHONUTF8=1"
set "TDD_PROJECT_DIR=%~dp0"
start "TDD" "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -Command "try { $ErrorActionPreference = 'Stop'; Set-Location -LiteralPath $env:TDD_PROJECT_DIR; if (-not (Test-Path -LiteralPath '.\.venv\Scripts\python.exe' -PathType Leaf)) { throw 'TDD is not set up yet: do step 3 of README.md first.' }; & '.\.venv\Scripts\python.exe' main.py; $code = $LASTEXITCODE; if ($code -ne 0) { Read-Host 'Press Enter to close' }; exit $code } catch { Write-Host $_ -ForegroundColor Red; Read-Host 'Press Enter to close'; exit 1 }"
exit /b %ERRORLEVEL%
