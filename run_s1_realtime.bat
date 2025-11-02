@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================================
echo 🔍 [S1] 실시간 주식 모니터링 시작
echo ========================================================
echo.
echo 모니터링 설정:
echo   - 거래일: 08:00 ~ 20:00
echo   - 간격: 60초 (1분)
echo   - Summary 탭의 종목만 모니터링
echo   - 매수선 접근 시 텔레그램 알림
echo.
echo 종료하려면 Ctrl+C를 누르세요.
echo ========================================================
echo.

python Real_Time_Monitor_S1.py ^
  --appkey IweTdkYa8JWDUOa8NohVSVeOiJ1THDGd_2x050A8XcU ^
  --secret eazu-jPNJpAsIVkaUTh3_88gUvXrCMJCwGF2AYRtBJs ^
  --interval 60

pause

