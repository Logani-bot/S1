@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ========================================
echo S1 작업 스케줄러 설정
echo ========================================
echo.
echo PowerShell 스크립트를 직접 실행합니다.
echo.
echo 관리자 권한이 필요할 수 있습니다.
echo.
echo 이 창은 닫지 마세요!
echo ========================================
echo.

powershell.exe -ExecutionPolicy Bypass -File "%~dp0setup_windows_scheduler_s1.ps1"

echo.
echo.
echo ========================================
echo 스크립트 실행 완료
echo ========================================
echo.
pause


