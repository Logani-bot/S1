@echo off
chcp 65001 > nul
echo ========================================
echo S1 작업 스케줄러 자동 설정
echo ========================================
echo.
echo 이 스크립트는 관리자 권한이 필요합니다.
echo.
echo 자동으로 PowerShell을 관리자 권한으로 실행합니다...
echo.
pause

powershell -ExecutionPolicy Bypass -Command "Start-Process powershell -Verb RunAs -ArgumentList '-NoExit -ExecutionPolicy Bypass -File \"%~dp0setup_windows_scheduler_s1.ps1\"'"

