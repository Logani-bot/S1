<<<<<<< HEAD
# S1 Trading System

시가총액 기반 주식 트레이딩 시스템

## 개요

S1 시스템은 **시가총액 1조 3천억원 이상** 종목을 대상으로 기술적 분석 기반 매매 신호를 생성하는 시스템입니다.

## 주요 구성 요소

### 1. Daily Market Cap Tracker (`Daily_MarketCap_Tracker.py`)
- 매일 시가총액 1조 3천억원 이상 종목 수집
- 키움 API `ka10099`를 사용하여 종목 리스트 조회
- 시가총액 계산: 상장주식수 × 전일종가
- 결과 저장: `output/marketcap_universe.xlsx`

### 2. Trading Signal System (`Trading_Signal_System_S1.py`)
- `marketcap_universe.xlsx`의 종목들을 분석
- 20일 이동평균선 기반 매수/매도 신호 생성
- 엔벨로프(-20%) 지지선 계산
- 3단계 매수선/매도선 계산
- 결과 저장: `output/trading_signals_s1.xlsx`

### 3. Real-Time Monitor (`Real_Time_Monitor_S1.py`)
- 실시간 가격 모니터링
- 매수/매도선 접근 시 텔레그램 알림
- 체결 시 즉시 알림

## 설치 및 실행

### 1. 패키지 설치
```bash
pip install -r requirements.txt
```

### 2. 환경 변수 설정
`.env` 파일 생성:
```
TELEGRAM_BOT_TOKEN=your_bot_token
TELEGRAM_CHAT_ID=your_chat_id
KIWOOM_APPKEY=your_appkey
KIWOOM_SECRET=your_secret
```

### 3. 일일 실행
```bash
RUN_S1_DAILY.bat
```
또는
```bash
python Daily_MarketCap_Tracker.py --appkey YOUR_APPKEY --secret YOUR_SECRET
python Trading_Signal_System_S1.py --appkey YOUR_APPKEY --secret YOUR_SECRET --alert-threshold 10.0
```

### 4. 실시간 모니터링
```bash
run_s1_realtime.bat
```

## 파일 구조

```
S1/
├── Daily_MarketCap_Tracker.py      # 시가총액 유니버스 수집
├── Trading_Signal_System_S1.py    # 매매 신호 생성
├── Real_Time_Monitor_S1.py         # 실시간 모니터링
├── telegram_notifier.py            # 텔레그램 알림
├── trading_day_utils.py            # 거래일 체크 유틸리티
├── contact_price_calculator.py    # 호가 계산 유틸리티
├── requirements.txt                # Python 패키지 목록
├── output/                         # 출력 파일
│   ├── marketcap_universe.xlsx
│   └── trading_signals_s1.xlsx
└── logs/                          # 로그 파일
```

## 알림 시스템

- **일일 리포트**: 매일 거래일 종료 후 전송 (접근 중인 종목 요약)
- **실시간 알림**: 매수/매도선 접근 시 즉시 전송
- **알림 접두사**: `[S1]` 사용

## S2 시스템과의 차이

- **S2**: 거래대금 상위 종목 기반 (`turnover_universe.xlsx`)
- **S1**: 시가총액 상위 종목 기반 (`marketcap_universe.xlsx`)
- 두 시스템은 독립적으로 운영되며, 별도의 텔레그램 알림을 제공합니다.

## 라이선스

Private

=======
# S1
S1 
>>>>>>> aef0ade7554fd192baa4e2ef3bb2af5994b53ff5
