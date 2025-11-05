@echo off
chcp 65001 >nul
cd /d "%~dp0"

REM 콘솔 창 크기 최대화
mode con: cols=120 lines=40

echo ========================================================
echo 🔍 [S1] 실시간 주식 모니터링 시작 (표시 모드)
echo ========================================================
echo.
echo 모니터링 설정:
echo   - 거래일: 08:00 ~ 20:00
echo   - 간격: 60초 (1분)
echo   - Summary 탭의 종목만 모니터링
echo   - 매수선 접근 시 텔레그램 알림
echo.
echo 이 창에서 로그를 실시간으로 확인할 수 있습니다.
echo.
echo ========================================================
echo.

REM 메인 모니터링 시작
python Real_Time_Monitor_S1.py ^
  --appkey IweTdkYa8JWDUOa8NohVSVeOiJ1THDGd_2x050A8XcU ^
  --secret eazu-jPNJpAsIVkaUTh3_88gUvXrCMJCwGF2AYRtBJs ^
  --interval 60

pause


