"""
Trading Day Utility Functions
거래일 체크 유틸리티 함수들

- 주말 체크 (토요일, 일요일)
- 공휴일 체크 (한국 공휴일)
- 거래일 체크 (평일 + 공휴일 제외)
"""

import holidays
from datetime import datetime, date, timedelta
from typing import Dict


def is_weekend(check_date: date = None) -> bool:
    """
    주말 체크 (토요일, 일요일)
    
    Args:
        check_date: 체크할 날짜 (None이면 오늘)
    
    Returns:
        True if 주말, False if 평일
    """
    if check_date is None:
        check_date = datetime.now().date()
    
    # weekday(): 월요일=0, 화요일=1, ..., 토요일=5, 일요일=6
    return check_date.weekday() >= 5


def is_holiday(check_date: date = None) -> bool:
    """
    공휴일 체크 (한국 공휴일)
    
    Args:
        check_date: 체크할 날짜 (None이면 오늘)
    
    Returns:
        True if 공휴일, False if 평일
    """
    if check_date is None:
        check_date = datetime.now().date()
    
    # 한국 공휴일 체크
    kr_holidays = holidays.SouthKorea()
    return check_date in kr_holidays


def is_trading_day(check_date: date = None) -> bool:
    """
    거래일 체크 (평일 + 공휴일 제외)
    
    Args:
        check_date: 체크할 날짜 (None이면 오늘)
    
    Returns:
        True if 거래일, False if 비거래일 (주말/공휴일)
    """
    if check_date is None:
        check_date = datetime.now().date()
    
    # 주말이면 거래일 아님
    if is_weekend(check_date):
        return False
    
    # 공휴일이면 거래일 아님
    if is_holiday(check_date):
        return False
    
    return True


def get_trading_day_info(check_date: date = None) -> dict:
    """
    거래일 상세 정보 반환
    
    Args:
        check_date: 체크할 날짜 (None이면 오늘)
    
    Returns:
        {
            'date': 날짜,
            'is_trading_day': 거래일 여부,
            'is_weekend': 주말 여부,
            'is_holiday': 공휴일 여부,
            'reason': 비거래일 사유 (거래일이 아닌 경우)
        }
    """
    if check_date is None:
        check_date = datetime.now().date()
    
    is_weekend_flag = is_weekend(check_date)
    is_holiday_flag = is_holiday(check_date)
    is_trading_day_flag = is_trading_day(check_date)
    
    reason = None
    if not is_trading_day_flag:
        if is_weekend_flag and is_holiday_flag:
            reason = "주말 + 공휴일"
        elif is_weekend_flag:
            reason = "주말"
        elif is_holiday_flag:
            reason = "공휴일"
    
    return {
        'date': check_date,
        'is_trading_day': is_trading_day_flag,
        'is_weekend': is_weekend_flag,
        'is_holiday': is_holiday_flag,
        'reason': reason
    }


def get_next_trading_day(start_date: date = None) -> date:
    """
    다음 거래일 반환
    
    Args:
        start_date: 시작 날짜 (None이면 오늘)
    
    Returns:
        다음 거래일
    """
    if start_date is None:
        start_date = datetime.now().date()
    
    next_date = start_date
    while True:
        next_date = datetime.combine(next_date, datetime.min.time()) + timedelta(days=1)
        next_date = next_date.date()
        
        if is_trading_day(next_date):
            return next_date


def get_previous_trading_day(start_date: date = None) -> date:
    """
    이전 거래일 반환 (최근 거래일)
    
    Args:
        start_date: 시작 날짜 (None이면 오늘)
    
    Returns:
        이전 거래일 (거래일이면 그대로 반환)
    """
    if start_date is None:
        start_date = datetime.now().date()
    
    prev_date = start_date
    while True:
        if is_trading_day(prev_date):
            return prev_date
        
        prev_date = datetime.combine(prev_date, datetime.min.time()) - timedelta(days=1)
        prev_date = prev_date.date()
        
        # 너무 과거로 가지 않도록 안전장치 (최대 7일 전까지만)
        if (start_date - prev_date).days > 7:
            return start_date


if __name__ == "__main__":
    # 테스트 코드
    print("=== 거래일 체크 테스트 ===")
    
    # 오늘 정보
    today_info = get_trading_day_info()
    print(f"오늘 ({today_info['date']}):")
    print(f"  거래일: {today_info['is_trading_day']}")
    print(f"  주말: {today_info['is_weekend']}")
    print(f"  공휴일: {today_info['is_holiday']}")
    if today_info['reason']:
        print(f"  사유: {today_info['reason']}")
    
    print()
    
    # 다음 거래일
    next_trading = get_next_trading_day()
    print(f"다음 거래일: {next_trading}")
    
    print()
    
    # 2025년 공휴일 예시
    print("=== 2025년 한국 공휴일 ===")
    kr_holidays = holidays.SouthKorea(years=2025)
    for holiday_date, holiday_name in sorted(kr_holidays.items()):
        print(f"{holiday_date}: {holiday_name}")
