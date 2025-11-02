"""
Real-time Stock Monitoring System

실시간 주식 모니터링 시스템 (거래일 08:00-20:00, 10분 간격)
- Summary 탭의 종목만 모니터링
- 현재가 기반 동적 20일선 계산
- 매수선 5% 이내 접근 시 알람
- 상태별 하루 1회 알람 (중복 방지)
- 주말/공휴일에는 모니터링 중단
"""

import sys
import logging
import requests
import pandas as pd
from datetime import datetime, time as time_type, timedelta
from pathlib import Path
import json
from typing import Dict, List, Tuple, Optional
import argparse
import time

# 거래일 체크 유틸리티 import
from trading_day_utils import is_trading_day, get_trading_day_info

# 로깅 설정
log_filename = f"realtime_monitor_{datetime.now().strftime('%Y%m%d')}.log"

# Windows 콘솔 인코딩 설정
import io
import sys
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s: %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.FileHandler(log_filename, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# 상수
SIGNAL_FILE = "output/trading_signals_s1.xlsx"  # S1 시스템 실시간 모니터링용 (Summary 탭)
ALERT_HISTORY_FILE = "alert_history.json"
MONITORING_START_TIME = time_type(8, 0)  # 08:00
MONITORING_END_TIME = time_type(20, 0)   # 20:00
DISTANCE_THRESHOLD = 5.0  # 5% 이내 접근 시 알람

# 키움 API 설정
KIWOOM_BASE_URL = "https://api.kiwoom.com"
KIWOOM_TOKEN_URL = "https://api.kiwoom.com/oauth2/token"
KIWOOM_TOKEN = None
APPKEY = None
SECRETKEY = None


def get_access_token(appkey: str, secretkey: str) -> Optional[str]:
    """
    키움 API 접근 토큰 발급
    """
    try:
        headers = {"Content-Type": "application/json;charset=UTF-8"}
        body = {
            "grant_type": "client_credentials",
            "appkey": appkey,
            "secretkey": secretkey
        }
        
        response = requests.post(KIWOOM_TOKEN_URL, headers=headers, json=body, timeout=20)
        response.raise_for_status()
        
        result = response.json()
        token = result.get("token") or result.get("access_token")
        
        if token:
            logger.info("✓ 접근 토큰 발급 성공")
            return token
        else:
            logger.error("✗ 접근 토큰 발급 실패")
            return None
    
    except Exception as e:
        logger.error(f"✗ 토큰 발급 중 오류: {e}")
        return None


def get_current_price(ticker: str, token: str) -> Optional[float]:
    """
    현재가 조회 (차트 API로 최신 데이터 조회)
    
    Args:
        ticker: 종목 코드
        token: 접근 토큰
    
    Returns:
        현재가 (실패 시 None)
    """
    try:
        url = f"{KIWOOM_BASE_URL}/api/dostk/chart"
        
        headers = {
            "authorization": f"Bearer {token}",
            "Content-Type": "application/json;charset=UTF-8",
            "api-id": "ka10081",
            "cont-yn": "N",
            "next-key": ""
        }
        
        # 오늘 날짜
        today = datetime.now().strftime("%Y%m%d")
        
        body = {
            "stk_cd": ticker,
            "base_dt": today,
            "upd_stkpc_tp": "1"
        }
        
        response = requests.post(url, headers=headers, json=body, timeout=10)
        response.raise_for_status()
        
        result = response.json()
        
        # 데이터 추출
        records = result.get("stk_dt_pole_chart_qry")
        
        if not records or len(records) == 0:
            logger.warning(f"⚠ {ticker}: 현재가 데이터 없음")
            return None
        
        # 가장 최근 데이터 (첫 번째 항목)
        latest = records[0]
        
        # 현재가 추출 (첫 번째 키가 cur_pric)
        # Note: 'cur_pric' in latest가 작동하지 않는 버그가 있어 직접 접근 사용
        all_keys = list(latest.keys())
        if len(all_keys) > 0:
            first_key = all_keys[0]  # cur_pric
            try:
                current_price = float(str(latest[first_key]).replace(",", ""))
                if current_price > 0:
                    return current_price
            except (ValueError, TypeError) as e:
                logger.warning(f"⚠ {ticker}: 현재가 파싱 실패 ({e})")
                return None
        
        logger.warning(f"⚠ {ticker}: 현재가 파싱 실패 (키 없음)")
        return None
    
    except Exception as e:
        logger.error(f"✗ {ticker} 현재가 조회 실패: {e}")
        return None


def get_enhanced_price_data(ticker: str, token: str) -> Optional[Dict]:
    """
    확장된 가격 데이터 조회 (현재가, 저가, 고가 포함)
    
    Args:
        ticker: 종목 코드
        token: 접근 토큰
    
    Returns:
        가격 데이터 딕셔너리 또는 None
    """
    try:
        url = f"{KIWOOM_BASE_URL}/api/dostk/chart"
        
        headers = {
            "authorization": f"Bearer {token}",
            "Content-Type": "application/json;charset=UTF-8",
            "api-id": "ka10081",
            "cont-yn": "N",
            "next-key": ""
        }
        
        # 오늘 날짜
        today = datetime.now().strftime("%Y%m%d")
        
        # KRX+NXT 통합 기준: 종목코드에 _AL 접미사 추가
        integrated_ticker = f"{ticker}_AL"
        
        body = {
            "stk_cd": integrated_ticker,  # 통합 종목코드 사용
            "base_dt": today,
            "upd_stkpc_tp": "1",  # 수정주가
            "stex_tp": "3"  # 통합 (KRX+NXT)
        }
        
        response = requests.post(url, headers=headers, json=body, timeout=10)
        response.raise_for_status()
        
        result = response.json()
        
        # 데이터 추출
        records = result.get("stk_dt_pole_chart_qry")
        
        if not records or len(records) == 0:
            logger.warning(f"⚠ {ticker}: 가격 데이터 없음")
            return None
        
        # 가장 최근 데이터 (첫 번째 항목)
        latest = records[0]
        
        # 명시적 키 이름으로 직접 접근 (올바른 매핑)
        data = {}
        data['current'] = float(str(latest.get('cur_prc', 0)).replace(",", ""))    # 현재가
        data['low'] = float(str(latest.get('low_pric', 0)).replace(",", ""))       # 저가
        data['high'] = float(str(latest.get('high_pric', 0)).replace(",", ""))     # 고가
        data['open'] = float(str(latest.get('open_pric', 0)).replace(",", ""))     # 시가
        data['volume'] = int(str(latest.get('trde_qty', 0)).replace(",", ""))      # 거래량
        
        # 필수 데이터 확인
        if 'current' not in data or data['current'] <= 0:
            logger.warning(f"⚠ {ticker}: 현재가 데이터 없음")
            return None
        
        return data
    
    except Exception as e:
        logger.error(f"✗ {ticker} 가격 데이터 조회 실패: {e}")
        return None


# fetch_chart_data 함수 제거 - Excel에서 매수선을 직접 읽어오므로 불필요


def calculate_tick_unit(price: float) -> int:
    """
    한국 주식시장 정확한 호가 단위 계산
    
    Args:
        price: 기준 가격
    
    Returns:
        호가 단위
    """
    if price < 2000:
        return 1
    elif price < 5000:
        return 5
    elif price < 20000:
        return 10
    elif price < 50000:
        return 50
    elif price < 200000:
        return 100
    elif price < 500000:
        return 500
    else:
        return 1000


def get_nearest_tick_price(price: float) -> float:
    """
    가장 가까운 정규 호가 가격 계산 (항상 윗 호가)
    
    Args:
        price: 기준 가격
    
    Returns:
        가장 가까운 정규 호가 가격 (항상 윗 호가)
    """
    tick_unit = calculate_tick_unit(price)
    
    # 현재 가격이 정확히 호가 단위에 맞는 경우
    if price % tick_unit == 0:
        return price
    
    # 현재 가격이 호가 단위 사이에 있는 경우 항상 윗 호가로 설정
    lower_tick = (price // tick_unit) * tick_unit
    upper_tick = lower_tick + tick_unit
    
    return upper_tick


def get_one_tick_up_price(price: float) -> float:
    """
    한 호가 위 가격 계산
    
    Args:
        price: 기준 가격
    
    Returns:
        한 호가 위 가격
    """
    nearest_tick = get_nearest_tick_price(price)
    tick_unit = calculate_tick_unit(nearest_tick)
    return nearest_tick + tick_unit


# solve_contact_price 함수 제거 - Excel 기반 시스템에서는 불필요


def calculate_monitoring_interval(current_price: float, envelope: float) -> int:
    """
    동적 모니터링 간격 계산
    
    Args:
        current_price: 현재가
        envelope: 엔벨로프 지지선
    
    Returns:
        모니터링 간격 (초)
    """
    if current_price is None or envelope is None or envelope == 0:
        return 600  # 기본 10분
    
    # 현재가와 엔벨로프 지지선 간의 거리 계산
    distance_pct = ((current_price - envelope) / envelope) * 100
    
    # 거리에 따른 간격 설정
    if distance_pct <= 1.0:  # 1% 이내
        return 60   # 1분
    elif distance_pct <= 3.0:  # 3% 이내
        return 180  # 3분
    elif distance_pct <= 10.0:  # 10% 이내
        return 600  # 10분
    else:  # 10% 이상
        return 1800  # 30분


def calculate_low_price_distance(low_price: float, target_price: float) -> float:
    """
    저가 기준 이격도 계산
    
    Args:
        low_price: 당일 저가
        target_price: 목표가 (매수선)
    
    Returns:
        이격도 (%)
    """
    if low_price is None or target_price is None or target_price == 0:
        return None
    
    return ((low_price - target_price) / target_price) * 100


# calculate_dynamic_ma20_and_buy_lines 함수 제거 - Excel에서 매수선을 직접 읽어오므로 불필요


def load_summary_stocks_with_buy_lines() -> pd.DataFrame:
    """
    Summary 탭에서 모니터링 대상 종목과 매수선 로드
    
    Returns:
        DataFrame with columns: 티커, 종목명, 매수상태, 1차매수선, 2차매수선, 3차매수선
    """
    try:
        if not Path(SIGNAL_FILE).exists():
            logger.warning(f"⚠ {SIGNAL_FILE} 파일이 없습니다.")
            return pd.DataFrame()
        
        df = pd.read_excel(SIGNAL_FILE, sheet_name="Summary")
        
        if df.empty:
            logger.info("ℹ Summary 탭에 종목이 없습니다.")
            return pd.DataFrame()
        
        # 필요한 컬럼만 선택
        required_columns = ['티커', '종목명', '매수상태', '1차매수선', '2차매수선', '3차매수선']
        
        # 컬럼 존재 확인
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            logger.error(f"✗ 필수 컬럼이 없습니다: {missing_columns}")
            return pd.DataFrame()
        
        # 필요한 컬럼만 추출
        df_filtered = df[required_columns].copy()
        
        # 매수선이 유효한 종목만 필터링
        df_filtered = df_filtered.dropna(subset=['1차매수선', '2차매수선', '3차매수선'])
        
        logger.info(f"[OK] Summary 탭에서 {len(df_filtered)}개 종목 로드 (매수선 포함)")
        return df_filtered
    
    except Exception as e:
        logger.error(f"✗ Summary 탭 로드 실패: {e}")
        return pd.DataFrame()


def load_alert_history() -> Dict:
    """
    알람 히스토리 로드 (오늘자)
    
    Returns:
        {
            "date": "2025-10-14",
            "alerts": {
                "005930": {
                    "READY_BUY1_5%": True,
                    "BOUGHT_1": True
                }
            }
        }
    """
    today = datetime.now().strftime("%Y-%m-%d")
    
    if not Path(ALERT_HISTORY_FILE).exists():
        return {
            "date": today,
            "alerts": {}
        }
    
    try:
        with open(ALERT_HISTORY_FILE, "r", encoding="utf-8") as f:
            history = json.load(f)
        
        # 날짜가 다르면 초기화
        if history.get("date") != today:
            return {
                "date": today,
                "alerts": {}
            }
        
        return history
    
    except Exception as e:
        logger.error(f"✗ 알람 히스토리 로드 실패: {e}")
        return {
            "date": today,
            "alerts": {}
        }


def save_alert_history(history: Dict):
    """
    알람 히스토리 저장
    """
    try:
        with open(ALERT_HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2)
    
    except Exception as e:
        logger.error(f"✗ 알람 히스토리 저장 실패: {e}")


def safe_float(value):
    """안전한 float 변환 함수"""
    if value is None or value == "":
        return None
    try:
        if isinstance(value, str):
            return float(value.replace(",", ""))
        return float(value)
    except (ValueError, TypeError):
        return None


def get_sell_prices_from_excel(ticker: str) -> dict:
    """Excel에서 해당 종목의 매도가 정보를 가져오는 함수"""
    try:
        import pandas as pd
        df = pd.read_excel('output/trading_signals_s1.xlsx', sheet_name='Summary', dtype={'티커': str})
        
        # 해당 티커의 행 찾기
        stock_row = df[df['티커'] == ticker]
        
        if len(stock_row) == 0:
            return {}
        
        row = stock_row.iloc[0]
        
        return {
            'sell1': safe_float(row.get('1차매도선(+3%)')),
            'sell2': safe_float(row.get('2차매도선(+5%)')),
            'sell3': safe_float(row.get('3차매도선(+7%)'))
        }
    except Exception as e:
        logger.warning(f"매도가 정보 조회 실패 ({ticker}): {e}")
        return {}


def calculate_low_price_distance(low_price: float, target_price: float) -> float:
    """
    저가 기준 이격도 계산 (부동소수점 오차 보정)
    
    Args:
        low_price: 저가
        target_price: 목표가 (매수선)
    
    Returns:
        이격도 (%)
    """
    if low_price is None or target_price is None or target_price == 0:
        return None
    
    # 부동소수점 오차 보정 (매우 작은 값은 0으로 처리)
    distance_pct = ((low_price - target_price) / target_price) * 100
    
    # 절댓값이 1e-10보다 작으면 0으로 처리 (극소값 제거)
    if abs(distance_pct) < 1e-10:
        return 0.0
    
    return distance_pct


def check_simplified_alert(
    ticker: str,
    stock_name: str,
    current_price: float,
    low_price: float,
    buy_status: str,
    buy1: float,
    buy2: float,
    buy3: float,
    history: Dict
) -> bool:
    """
    간단한 알람 조건 체크 및 텔레그램 전송 (Excel 기반)
    
    Args:
        ticker: 종목 코드
        stock_name: 종목명
        current_price: 현재가
        low_price: 저가
        buy_status: 매수 상태 (NONE, BOUGHT_1, BOUGHT_2, BOUGHT_3)
        buy1: 1차 매수선 (Excel에서 읽어온 값)
        buy2: 2차 매수선 (Excel에서 읽어온 값)
        buy3: 3차 매수선 (Excel에서 읽어온 값)
        history: 알람 히스토리
    
    Returns:
        알람 전송 여부
    """
    from telegram_notifier import send_realtime_alert
    
    # 매수선 값이 None인 경우 스킵
    if buy1 is None or buy2 is None or buy3 is None:
        logger.warning(f"⚠ {stock_name} ({ticker}): 매수선 데이터 없음 (buy1:{buy1}, buy2:{buy2}, buy3:{buy3})")
        return False
    
    alerts = history.get("alerts", {})
    ticker_alerts = alerts.get(ticker, {})
    
    # 1차 매수선 저가 기준 이격도 계산
    if buy_status == "NONE":
        # 저가 기준 이격도 계산
        low_dist_buy1 = calculate_low_price_distance(low_price, buy1)
        
        # 저가가 매수선에 도달한 경우 (마이너스 이격도) - 매수 체결!
        if low_dist_buy1 <= 0:
            alert_key = "BUY1_EXECUTED"
            alert_type = "1차 매수 체결!"
            
            if not ticker_alerts.get(alert_key, False):
                # 매도가 정보 가져오기
                sell_prices = get_sell_prices_from_excel(ticker)
                
                send_realtime_alert(
                    alert_type=alert_type,
                    stock_name=stock_name,
                    ticker=ticker,
                    current_price=current_price,
                    target_price=buy1,
                    distance_pct=low_dist_buy1,
                    recipients=["all"],
                    sell_prices=sell_prices,
                    system_label="S1"
                )
                
                ticker_alerts[alert_key] = True
                alerts[ticker] = ticker_alerts
                history["alerts"] = alerts
                save_alert_history(history)
                
                logger.info(f"🎯🎯 {stock_name} ({ticker}): 1차 매수 체결! 알람 전송")
                return True
        
        # 저가가 매수선 1% 이내 접근한 경우
        elif 0 < low_dist_buy1 <= 1.0:
            alert_key = "READY_BUY1_1%"
            alert_type = "1차 매수선 1% 인접"
            
            if not ticker_alerts.get(alert_key, False):
                send_realtime_alert(
                    alert_type=alert_type,
                    stock_name=stock_name,
                    ticker=ticker,
                    current_price=current_price,
                    target_price=buy1,
                    distance_pct=low_dist_buy1,
                    recipients=["all"],
                    system_label="S1"
                )
                
                ticker_alerts[alert_key] = True
                alerts[ticker] = ticker_alerts
                history["alerts"] = alerts
                save_alert_history(history)
                
                logger.info(f"🔴 {stock_name} ({ticker}): 1차 매수선 저가 기준 1% 인접 알람 전송")
                return True
        
        # 저가가 매수선 3% 이내 접근한 경우 (1% 알림이 없었던 경우만)
        elif 0 < low_dist_buy1 <= 3.0:
            alert_key = "READY_BUY1_3%"
            alert_type = "1차 매수선 3% 인접"
            
            if not ticker_alerts.get(alert_key, False) and not ticker_alerts.get("READY_BUY1_1%", False):
                send_realtime_alert(
                    alert_type=alert_type,
                    stock_name=stock_name,
                    ticker=ticker,
                    current_price=current_price,
                    target_price=buy1,
                    distance_pct=low_dist_buy1,
                    recipients=["all"],
                    system_label="S1"
                )
                
                ticker_alerts[alert_key] = True
                alerts[ticker] = ticker_alerts
                history["alerts"] = alerts
                save_alert_history(history)
                
                logger.info(f"🟠 {stock_name} ({ticker}): 1차 매수선 저가 기준 3% 인접 알람 전송")
                return True
        
        # 저가가 매수선 5% 이내 접근한 경우 (1%, 3% 알림이 없었던 경우만)
        elif 0 < low_dist_buy1 <= 5.0:
            alert_key = "READY_BUY1_5%"
            alert_type = "1차 매수선 5% 인접"
            
            if (not ticker_alerts.get(alert_key, False) and 
                not ticker_alerts.get("READY_BUY1_1%", False) and 
                not ticker_alerts.get("READY_BUY1_3%", False)):
                send_realtime_alert(
                    alert_type=alert_type,
                    stock_name=stock_name,
                    ticker=ticker,
                    current_price=current_price,
                    target_price=buy1,
                    distance_pct=low_dist_buy1,
                    recipients=["all"],
                    system_label="S1"
                )
                
                ticker_alerts[alert_key] = True
                alerts[ticker] = ticker_alerts
                history["alerts"] = alerts
                save_alert_history(history)
                
                logger.info(f"🟡 {stock_name} ({ticker}): 1차 매수선 저가 기준 5% 인접 알람 전송")
                return True
    
    # 2차 매수선 저가 기준 이격도 계산 (BOUGHT_1 상태일 때만)
    elif buy_status == "BOUGHT_1":
        low_dist_buy2 = calculate_low_price_distance(low_price, buy2)
        
        # 저가가 매수선에 도달한 경우 (마이너스 이격도) - 매수 체결!
        if low_dist_buy2 <= 0:
            alert_key = "BUY2_EXECUTED"
            alert_type = "2차 매수 체결!"
            
            if not ticker_alerts.get(alert_key, False):
                # 매도가 정보 가져오기
                sell_prices = get_sell_prices_from_excel(ticker)
                
                send_realtime_alert(
                    alert_type=alert_type,
                    stock_name=stock_name,
                    ticker=ticker,
                    current_price=current_price,
                    target_price=buy2,
                    distance_pct=low_dist_buy2,
                    recipients=["all"],
                    sell_prices=sell_prices,
                    system_label="S1"
                )
                
                ticker_alerts[alert_key] = True
                alerts[ticker] = ticker_alerts
                history["alerts"] = alerts
                save_alert_history(history)
                
                logger.info(f"🎯🎯 {stock_name} ({ticker}): 2차 매수 체결! 알람 전송")
                return True
        
        # 저가가 매수선 1% 이내 접근한 경우
        elif 0 < low_dist_buy2 <= 1.0:
            alert_key = "READY_BUY2_1%"
            alert_type = "2차 매수선 1% 인접"
            
            if not ticker_alerts.get(alert_key, False):
                send_realtime_alert(
                    alert_type=alert_type,
                    stock_name=stock_name,
                    ticker=ticker,
                    current_price=current_price,
                    low_price=low_price,
                    target_price=buy2,
                    distance_pct=low_dist_buy2,
                    recipients=["all"],
                    system_label="S1"
                )
                
                ticker_alerts[alert_key] = True
                alerts[ticker] = ticker_alerts
                history["alerts"] = alerts
                save_alert_history(history)
                
                logger.info(f"🔴 {stock_name} ({ticker}): 2차 매수선 저가 기준 1% 인접 알람 전송")
                return True
        
        # 저가가 매수선 3% 이내 접근한 경우 (1% 알림이 없었던 경우만)
        elif 0 < low_dist_buy2 <= 3.0:
            alert_key = "READY_BUY2_3%"
            alert_type = "2차 매수선 3% 인접"
            
            if not ticker_alerts.get(alert_key, False) and not ticker_alerts.get("READY_BUY2_1%", False):
                send_realtime_alert(
                    alert_type=alert_type,
                    stock_name=stock_name,
                    ticker=ticker,
                    current_price=current_price,
                    low_price=low_price,
                    target_price=buy2,
                    distance_pct=low_dist_buy2,
                    recipients=["all"],
                    system_label="S1"
                )
                
                ticker_alerts[alert_key] = True
                alerts[ticker] = ticker_alerts
                history["alerts"] = alerts
                save_alert_history(history)
                
                logger.info(f"🟠 {stock_name} ({ticker}): 2차 매수선 저가 기준 3% 인접 알람 전송")
                return True
        
        # 저가가 매수선 5% 이내 접근한 경우 (1%, 3% 알림이 없었던 경우만)
        elif 0 < low_dist_buy2 <= 5.0:
            alert_key = "READY_BUY2_5%"
            alert_type = "2차 매수선 5% 인접"
            
            if (not ticker_alerts.get(alert_key, False) and 
                not ticker_alerts.get("READY_BUY2_1%", False) and 
                not ticker_alerts.get("READY_BUY2_3%", False)):
                send_realtime_alert(
                    alert_type=alert_type,
                    stock_name=stock_name,
                    ticker=ticker,
                    current_price=current_price,
                    low_price=low_price,
                    target_price=buy2,
                    distance_pct=low_dist_buy2,
                    recipients=["all"],
                    system_label="S1"
                )
                
                ticker_alerts[alert_key] = True
                alerts[ticker] = ticker_alerts
                history["alerts"] = alerts
                save_alert_history(history)
                
                logger.info(f"🟡 {stock_name} ({ticker}): 2차 매수선 저가 기준 5% 인접 알람 전송")
                return True
    
    # 3차 매수선 저가 기준 이격도 계산 (BOUGHT_2 상태일 때만)
    elif buy_status == "BOUGHT_2":
        low_dist_buy3 = calculate_low_price_distance(low_price, buy3)
        
        # 저가가 매수선에 도달한 경우 (마이너스 이격도) - 매수 체결!
        if low_dist_buy3 <= 0:
            alert_key = "BUY3_EXECUTED"
            alert_type = "3차 매수 체결!"
            
            if not ticker_alerts.get(alert_key, False):
                # 매도가 정보 가져오기
                sell_prices = get_sell_prices_from_excel(ticker)
                
                send_realtime_alert(
                    alert_type=alert_type,
                    stock_name=stock_name,
                    ticker=ticker,
                    current_price=current_price,
                    target_price=buy3,
                    distance_pct=low_dist_buy3,
                    recipients=["all"],
                    sell_prices=sell_prices,
                    system_label="S1"
                )
                
                ticker_alerts[alert_key] = True
                alerts[ticker] = ticker_alerts
                history["alerts"] = alerts
                save_alert_history(history)
                
                logger.info(f"🎯🎯🎯 {stock_name} ({ticker}): 3차 매수 체결! 알람 전송")
                return True
        
        # 저가가 매수선 1% 이내 접근한 경우
        elif 0 < low_dist_buy3 <= 1.0:
            alert_key = "READY_BUY3_1%"
            alert_type = "3차 매수선 1% 인접"
            
            if not ticker_alerts.get(alert_key, False):
                send_realtime_alert(
                    alert_type=alert_type,
                    stock_name=stock_name,
                    ticker=ticker,
                    current_price=current_price,
                    low_price=low_price,
                    target_price=buy3,
                    distance_pct=low_dist_buy3,
                    recipients=["all"],
                    system_label="S1"
                )
                
                ticker_alerts[alert_key] = True
                alerts[ticker] = ticker_alerts
                history["alerts"] = alerts
                save_alert_history(history)
                
                logger.info(f"🔴 {stock_name} ({ticker}): 3차 매수선 저가 기준 1% 인접 알람 전송")
                return True
        
        # 저가가 매수선 3% 이내 접근한 경우 (1% 알림이 없었던 경우만)
        elif 0 < low_dist_buy3 <= 3.0:
            alert_key = "READY_BUY3_3%"
            alert_type = "3차 매수선 3% 인접"
            
            if not ticker_alerts.get(alert_key, False) and not ticker_alerts.get("READY_BUY3_1%", False):
                send_realtime_alert(
                    alert_type=alert_type,
                    stock_name=stock_name,
                    ticker=ticker,
                    current_price=current_price,
                    low_price=low_price,
                    target_price=buy3,
                    distance_pct=low_dist_buy3,
                    recipients=["all"],
                    system_label="S1"
                )
                
                ticker_alerts[alert_key] = True
                alerts[ticker] = ticker_alerts
                history["alerts"] = alerts
                save_alert_history(history)
                
                logger.info(f"🟠 {stock_name} ({ticker}): 3차 매수선 저가 기준 3% 인접 알람 전송")
                return True
        
        # 저가가 매수선 5% 이내 접근한 경우 (1%, 3% 알림이 없었던 경우만)
        elif 0 < low_dist_buy3 <= 5.0:
            alert_key = "READY_BUY3_5%"
            alert_type = "3차 매수선 5% 인접"
            
            if (not ticker_alerts.get(alert_key, False) and 
                not ticker_alerts.get("READY_BUY3_1%", False) and 
                not ticker_alerts.get("READY_BUY3_3%", False)):
                send_realtime_alert(
                    alert_type=alert_type,
                    stock_name=stock_name,
                    ticker=ticker,
                    current_price=current_price,
                    low_price=low_price,
                    target_price=buy3,
                    distance_pct=low_dist_buy3,
                    recipients=["all"],
                    system_label="S1"
                )
                
                ticker_alerts[alert_key] = True
                alerts[ticker] = ticker_alerts
                history["alerts"] = alerts
                save_alert_history(history)
                
                logger.info(f"🟡 {stock_name} ({ticker}): 3차 매수선 저가 기준 5% 인접 알람 전송")
                return True
    
    return False


def is_monitoring_time(force_mode: bool = False) -> bool:
    """
    모니터링 시간대 체크 (거래일 08:00-20:00)
    
    Args:
        force_mode: 강제 실행 모드 (거래일 체크 무시)
    """
    now = datetime.now()
    
    # 거래일 체크 (강제 실행 모드가 아닌 경우에만)
    if not force_mode:
        if not is_trading_day(now.date()):
            return False
    
    # 시간 체크 (08:00-20:00)
    return MONITORING_START_TIME <= now.time() <= MONITORING_END_TIME


def run_simplified_monitoring_cycle():
    """
    단순화된 모니터링 사이클 실행 (Excel 기반)
    """
    global KIWOOM_TOKEN
    
    try:
        # 1. 접근 토큰 발급 (또는 재사용)
        if not KIWOOM_TOKEN:
            KIWOOM_TOKEN = get_access_token(APPKEY, SECRETKEY)
            if not KIWOOM_TOKEN:
                logger.error("✗ 토큰 발급 실패")
                return False
        
        # 2. Excel에서 종목과 매수선 로드
        df_summary = load_summary_stocks_with_buy_lines()
        if df_summary.empty:
            logger.info("ℹ 모니터링 대상 종목이 없습니다.")
            return True
        
        # 3. 알람 히스토리 로드
        alert_history = load_alert_history()
        
        # 4. 각 종목별 모니터링
        current_time = datetime.now()
        alert_count = 0
        checked_count = 0
        
        for idx, row in df_summary.iterrows():
            ticker = str(row.get("티커", "")).zfill(6)
            stock_name = row.get("종목명", "")
            buy_status = row.get("매수상태", "NONE")
            
            # Excel에서 매수선 읽기 (안전한 변환)
            buy1_raw = row.get("1차매수선", 0)
            buy2_raw = row.get("2차매수선", 0)
            buy3_raw = row.get("3차매수선", 0)
            
            # 문자열을 숫자로 안전하게 변환
            def safe_float(value):
                if value is None or value == "":
                    return None
                try:
                    if isinstance(value, str):
                        return float(value.replace(",", ""))
                    return float(value)
                except (ValueError, TypeError):
                    return None
            
            buy1 = safe_float(buy1_raw)
            buy2 = safe_float(buy2_raw)
            buy3 = safe_float(buy3_raw)
            
            checked_count += 1
            logger.info(f"\n[{checked_count}] {stock_name} ({ticker}) 모니터링 중...")
            
            # API 호출 제한 방지 (0.5초 대기)
            if checked_count > 1:
                time.sleep(0.5)
            
            # 확장된 가격 데이터 조회 (현재가, 저가만)
            price_data = get_enhanced_price_data(ticker, KIWOOM_TOKEN)
            if not price_data:
                logger.warning(f"⚠ {stock_name}: 가격 데이터 조회 실패, 스킵")
                continue
            
            current_price = price_data.get('current', 0)
            low_price = price_data.get('low', 0)
            high_price = price_data.get('high', 0)
            
            logger.info(f"  [현재가] 현재가: {current_price:,.0f}원")
            logger.info(f"  [저가] 저가: {low_price:,.0f}원")
            logger.info(f"  [고가] 고가: {high_price:,.0f}원")
            
            # Excel에서 읽어온 매수선 표시
            logger.info(f"  [매수선] 1차: {buy1:,.0f}원, 2차: {buy2:,.0f}원, 3차: {buy3:,.0f}원")
            
            # 현재가 기준 이격도 계산
            dist1 = ((current_price - buy1) / buy1) * 100 if buy1 and buy1 > 0 else None
            dist2 = ((current_price - buy2) / buy2) * 100 if buy2 and buy2 > 0 else None
            dist3 = ((current_price - buy3) / buy3) * 100 if buy3 and buy3 > 0 else None
            
            logger.info(f"  [이격도] 1차: {dist1:.1f}%, 2차: {dist2:.1f}%, 3차: {dist3:.1f}%")
            
            # 저가 기준 인접 알림 체크
            if check_simplified_alert(ticker, stock_name, current_price, low_price, buy_status, buy1, buy2, buy3, alert_history):
                alert_count += 1
        
        logger.info("\n" + "=" * 80)
        logger.info("[OK] 단순화된 모니터링 사이클 완료")
        logger.info(f"  전체 종목: {len(df_summary)}개")
        logger.info(f"  체크한 종목: {checked_count}개")
        logger.info(f"  전송 알람: {alert_count}개")
        logger.info("=" * 80)
        
        return True
    
    except Exception as e:
        logger.error(f"[ERROR] 시스템 오류: {e}")
        
        try:
            from telegram_notifier import send_error_alert
            send_error_alert(str(e), "Real_Time_Monitor", recipients=["me"])  # 에러는 본인만
        except:
            pass
        
        return False


# run_monitoring_cycle 함수 제거 - 단순화된 버전으로 교체됨


def main():
    """
    메인 함수 - 단순화된 실시간 모니터링 (Excel 기반)
    """
    global APPKEY, SECRETKEY, KIWOOM_TOKEN
    
    # 인자 파싱
    parser = argparse.ArgumentParser(description="실시간 주식 모니터링 (Excel 기반)")
    parser.add_argument("--appkey", required=True, help="키움 APPKEY")
    parser.add_argument("--secret", required=True, help="키움 SECRETKEY")
    parser.add_argument("--interval", type=int, default=60, help="모니터링 간격 (초, 기본값: 60)")
    parser.add_argument("--force", action="store_true", help="거래일 체크 무시하고 강제 실행")
    args = parser.parse_args()
    
    APPKEY = args.appkey
    SECRETKEY = args.secret
    base_interval = args.interval
    
    logger.info("=" * 80)
    logger.info("🔍 실시간 주식 모니터링 시작 (Excel 기반 단순화 버전)")
    logger.info(f"⏰ 모니터링 간격: {base_interval}초 ({base_interval//60}분)")
    logger.info("📊 Excel에서 매수선 읽어오기")
    logger.info("🎯 저가 기준 터치 감지 활성화")
    logger.info("🕐 모니터링 시간: 거래일 08:00-20:00")
    logger.info("📅 주말/공휴일 모니터링 중단")
    logger.info("=" * 80)
    
    cycle_count = 0
    
    while True:
        cycle_count += 1
        current_time = datetime.now()
        
        # 모니터링 시간대 체크 (강제 실행 옵션이 없는 경우에만)
        if not args.force:
            if not is_monitoring_time():
                # 거래일 정보 가져오기
                trading_info = get_trading_day_info()
                if not trading_info['is_trading_day']:
                    logger.info(f"\n[사이클 {cycle_count}] 비거래일입니다 ({trading_info['reason']})")
                else:
                    logger.info(f"\n[사이클 {cycle_count}] 모니터링 시간대가 아닙니다 (거래일 08:00-20:00)")
                
                logger.info(f"⏰ {base_interval}초 후 재확인...")
                logger.info("강제 실행하려면 --force 옵션을 사용하세요.")
                time.sleep(base_interval)
                continue
        else:
            # 강제 실행 모드에서는 시간대만 체크
            if not is_monitoring_time(force_mode=True):
                logger.info(f"\n[사이클 {cycle_count}] 모니터링 시간대가 아닙니다 (08:00-20:00)")
                logger.info(f"⏰ {base_interval}초 후 재확인...")
                time.sleep(base_interval)
                continue
        
        logger.info(f"\n{'=' * 80}")
        logger.info(f"[사이클 {cycle_count}] {current_time.strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info(f"{'=' * 80}")
        
        # 단순화된 모니터링 실행
        success = run_simplified_monitoring_cycle()
        
        if not success:
            logger.warning("[WARNING] 모니터링 실패, 재시도...")
        
        # 다음 실행까지 대기
        logger.info(f"\n[간격] {base_interval}초 후 다음 사이클 실행...")
        logger.info(f"   종료하려면 Ctrl+C를 누르세요.")
        
        try:
            time.sleep(base_interval)
        except KeyboardInterrupt:
            logger.info("\n" + "=" * 80)
            logger.info("[STOP] 사용자가 모니터링을 중지했습니다.")
            logger.info("=" * 80)
            break


if __name__ == "__main__":
    main()

