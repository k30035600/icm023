@echo off
REM PowerShell 스크립트를 관리자 권한 없이 우회 실행합니다.
chcp 65001 >nul
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -File "%~dp0run.ps1"
pause
