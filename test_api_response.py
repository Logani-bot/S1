#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""API 응답 구조 확인 테스트"""

import requests
import json
from datetime import date

def get_access_token(appkey: str, secret: str):
    """키움 API 토큰 획득"""
    url = "https://api.kiwoom.com/oauth2/token"
    headers = {
        "Content-Type": "application/json"
    }
    body = {
        "grant_type": "client_credentials",
        "appkey": appkey,
        "appsecret": secret
    }
    
    response = requests.post(url, headers=headers, json=body, timeout=20)
    response.raise_for_status()
    result = response.json()
    
    print(f"토큰 응답: {json.dumps(result, indent=2, ensure_ascii=False)}")
    
    if result.get("access_token"):
        return result["access_token"]
    raise RuntimeError(f"토큰 획득 실패: {result}")

def test_market_cap_api(token: str, query_date: date):
    """시가총액 순위 API 테스트"""
    headers = {
        "authorization": f"Bearer {token}",
        "Content-Type": "application/json;charset=UTF-8",
        "api-id": "ka10032",
        "cont-yn": "N",
        "next-key": ""
    }
    
    body = {
        "qry_tp": "1",  # 시가총액 상위
        "rank_strt": "0",
        "rank_end": "10",  # 상위 10개만 테스트
        "stex_tp": "3",    # 통합
        "mang_stk_incls": "1",
        "mrkt_tp": "001",  # KRX
        "bas_dd": query_date.strftime("%Y%m%d")
    }
    
    url = "https://api.kiwoom.com/api/dostk/rkinfo"
    
    response = requests.post(url, headers=headers, json=body, timeout=20)
    response.raise_for_status()
    result = response.json()
    
    print("=" * 60)
    print(f"조회 날짜: {query_date}")
    print(f"API 응답 키: {list(result.keys())}")
    print(f"응답 전체:")
    print(json.dumps(result, indent=2, ensure_ascii=False))
    print("=" * 60)
    
    return result

if __name__ == "__main__":
    appkey = "IweTdkYa8JWDUOa8NohVSVeOiJ1THDGd_2x050A8XcU"
    secret = "eazu-jPNJpAsIVkaUTh3_88gUvXrCMJCwGF2AYRtBJs"
    
    # 최근 거래일 계산
    from trading_day_utils import get_previous_trading_day
    query_date = get_previous_trading_day()
    
    print(f"최근 거래일: {query_date}")
    
    token = get_access_token(appkey, secret)
    test_market_cap_api(token, query_date)

