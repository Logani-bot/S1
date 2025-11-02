@echo off
chcp 65001 > nul
REM ========================================
REM S1 Trading System - Daily Report
REM ========================================
REM 이 파일을 더블클릭하면 일일 리포트가 실행됩니다
REM 20:15에 자동 실행되는 것과 동일합니다
REM ========================================

echo ========================================
echo [S1] Trading System - Daily Report
echo ========================================
echo.

cd /d "%~dp0"

REM Step 1: 시가총액 기준 유니버스 업데이트
echo [1/2] marketcap_universe.xlsx 업데이트 중...
"C:\Program Files (x86)\Python311\python.exe" Daily_MarketCap_Tracker.py --appkey IweTdkYa8JWDUOa8NohVSVeOiJ1THDGd_2x050A8XcU --secret eazu-jPNJpAsIVkaUTh3_88gUvXrCMJCwGF2AYRtBJs

if %ERRORLEVEL% neq 0 (
    echo ERROR: Daily Market Cap Tracker failed!
    pause
    exit /b 1
)

echo.
echo ========================================
echo [2/2] trading_signals_s1.xlsx 업데이트 중...
"C:\Program Files (x86)\Python311\python.exe" Trading_Signal_System_S1.py --appkey IweTdkYa8JWDUOa8NohVSVeOiJ1THDGd_2x050A8XcU --secret eazu-jPNJpAsIVkaUTh3_88gUvXrCMJCwGF2AYRtBJs --alert-threshold 10.0

if %ERRORLEVEL% neq 0 (
    echo ERROR: Trading Signal System S1 failed!
    pause
    exit /b 1
)

echo.
echo ========================================
echo [S1] 업데이트 완료!
echo ========================================
echo.
echo output\marketcap_universe.xlsx
echo output\trading_signals_s1.xlsx
echo.

pause

