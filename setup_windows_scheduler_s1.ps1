# Windows 작업 스케줄러 자동 설정 스크립트
# S1 트레이딩 시스템용

# 인코딩 설정 (한글 출력을 위해)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['*:Encoding'] = 'utf8'

# PowerShell 관리자 권한 체크 (경고만 표시, 계속 진행)
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Host "⚠️  관리자 권한이 없습니다. 현재 사용자 권한으로 시도합니다." -ForegroundColor Yellow
    Write-Host "   만약 작업 등록에 실패하면, 관리자 권한으로 PowerShell을 실행해주세요." -ForegroundColor Yellow
    Write-Host ""
    Start-Sleep -Seconds 2
}

# 현재 스크립트 위치
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$batFile1 = Join-Path $scriptDir "RUN_S1_DAILY.bat"
$batFile2 = Join-Path $scriptDir "run_s1_realtime.bat"

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Windows 작업 스케줄러 설정 시작 (S1)" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# 작업 스케줄러 서비스 체크
try {
    $schedule = New-Object -ComObject Schedule.Service
    $schedule.Connect()
} catch {
    Write-Host "❌ 작업 스케줄러에 연결할 수 없습니다." -ForegroundColor Red
    exit 1
}

# 기존 작업 삭제 (있다면)
$taskNames = @("S1_Daily_Trading_Signal", "S1_Realtime_Monitor")
foreach ($taskName in $taskNames) {
    try {
        $schedule.GetFolder("\").DeleteTask($taskName, 0)
        Write-Host "✓ 기존 작업 삭제: $taskName" -ForegroundColor Yellow
    } catch {
        # 작업이 없으면 에러 무시
    }
}

# 작업 1: Daily Market Cap + Signal (매일 20:15)
Write-Host "`n작업 1 생성 중..." -ForegroundColor Green
$action1 = New-ScheduledTaskAction -Execute $batFile1
$trigger1 = New-ScheduledTaskTrigger -Daily -At "20:15"
$settings1 = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
Register-ScheduledTask -TaskName "S1_Daily_Trading_Signal" -Action $action1 -Trigger $trigger1 -Settings $settings1 -Description "S1 매일 시가총액 + 시그널 분석 (20:15)" -Force | Out-Null
Write-Host "✓ 작업 1 생성 완료: S1_Daily_Trading_Signal (매일 20:15)" -ForegroundColor Green

# 작업 2: Real-Time Monitor (평일 08:00) - 표시 모드
Write-Host "`n작업 2 생성 중... (표시 모드)" -ForegroundColor Green
$action2 = New-ScheduledTaskAction -Execute $batFile2
$trigger2 = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At "08:00"
$settings2 = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
$task2 = Register-ScheduledTask -TaskName "S1_Realtime_Monitor" -Action $action2 -Trigger $trigger2 -Settings $settings2 -Description "S1 실시간 주식 모니터링 (평일 08:00)" -Force

# 표시 모드로 설정
$task2.Principal.DisplayName = "S1 실시간 모니터"
$task2.Settings.Hidden = $false
$task2 | Set-ScheduledTask

Write-Host "✓ 작업 2 생성 완료: S1_Realtime_Monitor (평일 08:00, 표시)" -ForegroundColor Green

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "✅ 작업 스케줄러 설정 완료! (S1)" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Host "생성된 작업:" -ForegroundColor Yellow
Write-Host "  1. S1_Daily_Trading_Signal - 매일 20:15 실행 (백그라운드)" -ForegroundColor White
Write-Host "  2. S1_Realtime_Monitor - 평일 08:00 실행 (표시)" -ForegroundColor White

Write-Host "`n표시 모드 특징:" -ForegroundColor Yellow
Write-Host "  - 실시간 모니터링은 화면에 표시됩니다" -ForegroundColor White
Write-Host "  - 일일 리포트는 백그라운드에서 실행됩니다" -ForegroundColor White

Write-Host "`n테스트 방법:" -ForegroundColor Yellow
Write-Host "  1. 작업 스케줄러 열기: Win+R → taskschd.msc" -ForegroundColor White
Write-Host "  2. 작업 목록에서 S1_Realtime_Monitor 우클릭 → '실행'" -ForegroundColor White
Write-Host "  3. 새 창이 열리며 로그가 표시됩니다" -ForegroundColor White
Write-Host "`n" -ForegroundColor White

Write-Host "아무 키나 누르면 창이 닫힙니다..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")


