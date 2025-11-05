# 🎯 S1 시스템 24/7 자동 실행 설정 가이드

## ✅ 설정 완료 상태

S1 트레이딩 시스템이 24/7로 자동 실행되도록 설정되었습니다.

## 📋 생성된 파일

1. **setup_windows_scheduler_s1.ps1** - Windows 작업 스케줄러 자동 설정 스크립트
2. **run_s1_realtime_with_display.bat** - 표시 모드 실시간 모니터링 실행 파일

## 🚀 설정 방법

### 1단계: PowerShell 관리자 권한으로 실행

1. Windows 키 + X → "Windows PowerShell (관리자)" 선택
2. 또는 PowerShell을 관리자 권한으로 실행

### 2단계: 스케줄러 설정 스크립트 실행

```powershell
cd C:\Users\log\Desktop\Code\S1
.\setup_windows_scheduler_s1.ps1
```

### 3단계: 설정 확인

스크립트가 다음 작업들을 생성합니다:
- ✅ **S1_Daily_Trading_Signal** - 매일 20:15 실행
- ✅ **S1_Realtime_Monitor** - 평일 08:00 실행

## 📅 자동 실행 일정

### 일일 작업 (매일 20:15)
- **작업 이름**: S1_Daily_Trading_Signal
- **실행 파일**: RUN_S1_DAILY.bat
- **작업 내용**:
  1. 시가총액 유니버스 수집 (Daily_MarketCap_Tracker.py)
  2. 매매 신호 생성 (Trading_Signal_System_S1.py)
- **실행 모드**: 백그라운드 (숨김)

### 실시간 모니터링 (평일 08:00)
- **작업 이름**: S1_Realtime_Monitor
- **실행 파일**: run_s1_realtime.bat
- **작업 내용**:
  - 거래일 08:00-20:00 동안 실시간 모니터링
  - 매수선 접근 시 텔레그램 알림
  - 60초 간격으로 체크
- **실행 모드**: 표시 모드 (화면에 표시)

## 🔄 자동 재시작 설정

모든 작업은 다음과 같이 설정되어 있습니다:
- ✅ **배터리 모드에서도 실행**: AllowStartIfOnBatteries
- ✅ **배터리 절약 모드에서도 계속 실행**: DontStopIfGoingOnBatteries
- ✅ **오류 시 자동 재시작**: 최대 3회, 1분 간격
- ✅ **작업 시작 가능 시 즉시 실행**: StartWhenAvailable

## 📊 모니터링 시간

### 거래일 체크
- 주말 자동 스킵
- 공휴일 자동 스킵
- 거래일만 모니터링 실행

### 모니터링 시간대
- **시작**: 08:00
- **종료**: 20:00
- **간격**: 60초 (1분)

## 🧪 테스트 방법

### 작업 스케줄러에서 수동 실행
1. `Win + R` → `taskschd.msc` 입력
2. 작업 목록에서 작업 선택
3. 우클릭 → "실행" 선택
4. 로그 확인

### 배치 파일로 직접 실행
```batch
# 일일 리포트 수동 실행
RUN_S1_DAILY.bat

# 실시간 모니터링 수동 실행
run_s1_realtime.bat
```

## 📝 로그 확인

### 로그 파일 위치
- **일일 리포트**: `logs/s12_daily_YYYYMMDD.log` (없을 수 있음)
- **실시간 모니터링**: `realtime_monitor_YYYYMMDD.log`

### 로그 확인 방법
```powershell
# 최신 로그 파일 확인
Get-Content realtime_monitor_$(Get-Date -Format "yyyyMMdd").log -Tail 50

# 실시간 로그 모니터링
Get-Content realtime_monitor_$(Get-Date -Format "yyyyMMdd").log -Wait -Tail 20
```

## ⚠️ 주의사항

1. **관리자 권한 필요**: 스케줄러 설정 시 관리자 권한이 필요합니다.
2. **컴퓨터 켜져 있어야 함**: 자동 실행을 위해서는 컴퓨터가 켜져 있어야 합니다.
3. **인터넷 연결 필요**: 키움 API 접근을 위해 인터넷 연결이 필요합니다.

## 🔧 문제 해결

### 작업이 실행되지 않는 경우
1. 작업 스케줄러 서비스가 실행 중인지 확인
2. 작업의 "마지막 실행 결과" 확인
3. 로그 파일 확인

### 모니터링이 중단되는 경우
1. 컴퓨터 절전 모드 해제
2. 작업 스케줄러의 "작업이 실행 중이면 다음 규칙 적용 전에 완료 대기" 설정 확인

## ✅ 완료 확인

설정이 완료되면:
- ✅ 매일 20:15에 일일 리포트 자동 실행
- ✅ 평일 08:00에 실시간 모니터링 자동 시작
- ✅ 20:00에 자동 종료
- ✅ 다음 거래일 자동 재시작
- ✅ 비거래일 자동 스킵

**24/7 자동 감시 시스템이 준비되었습니다!** 🎉


