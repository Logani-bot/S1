#!/usr/bin/env python
# -*- coding: utf-8 -*-

from slack_notifier import send_slack_daily_report

# 테스트 데이터
test_alerts = [
    {
        '종목명': '테스트종목',
        '종가': 50000,
        '1차매수선(익일)': 45000,
        '1차매수선이격도(%)': 11.1,
        '알람상태': 'READY_BUY1'
    }
]

print("Slack 일일 리포트 테스트 전송 중...")
if send_slack_daily_report(test_alerts, 1, 'S1'):
    print("Slack 일일 리포트 전송 성공!")
else:
    print("Slack 일일 리포트 전송 실패")


