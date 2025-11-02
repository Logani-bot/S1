"""
실시간 감시용 "접점 가격" 산출 시스템
- 과거 19일 종가 합 S19 기반 접점 가격 계산
- KRX 호가 단위 정확 적용
- 고정점(Fixed Point) 수학적 해법
"""

def get_tick_size(p: float) -> int:
    """한국 주식시장 호가 단위 반환"""
    if p < 2000:
        return 1
    elif p < 5000:
        return 5
    elif p < 20000:
        return 10
    elif p < 50000:
        return 50
    elif p < 200000:
        return 100
    elif p < 500000:
        return 500
    else:
        return 1000


def ceil_tick(y: float) -> int:
    """가격을 해당 구간의 호가 단위에 맞춰 항상 윗 호가로 올림"""
    d = get_tick_size(y)
    q = int(y // d)
    return q * d if (q * d == y) else (q + 1) * d


def solve_contact_price(S19: float) -> int:
    """
    접점 가격 계산 (고정점 해법)
    
    Args:
        S19: 과거 19일 종가 합계
    
    Returns:
        접점 가격 p (정상 호가)
    """
    # 이론적 접점 계산
    x_star = S19 / 24.0
    
    # 항상 ceil_tick(x_star)에서 시작
    p = ceil_tick(x_star)
    
    # 반복 수행으로 실제 접점 찾기
    while True:
        d = get_tick_size(p)
        upper = (S19 + 25 * d) / 24.0
        
        # 접점 조건 검사: S19/24 <= p < (S19 + 25*delta)/24
        if x_star <= p < upper:
            return int(p)
        
        # 다음 호가로 이동
        p += d


def solve_contact_price_debug(S19: float) -> dict:
    """
    접점 가격 계산 (디버그 정보 포함)
    
    Args:
        S19: 과거 19일 종가 합계
    
    Returns:
        {
            "contact_price": 접점 가격,
            "x_star": 이론적 접점,
            "delta": 호가 단위,
            "upper": 상한값,
            "iterations": 반복 횟수
        }
    """
    x_star = S19 / 24.0
    p = ceil_tick(x_star)
    iterations = 0
    
    while True:
        iterations += 1
        d = get_tick_size(p)
        upper = (S19 + 25 * d) / 24.0
        
        if x_star <= p < upper:
            return {
                "contact_price": int(p),
                "x_star": x_star,
                "delta": d,
                "upper": upper,
                "iterations": iterations
            }
        
        p += d


def verify_contact_price(S19: float, p: int) -> bool:
    """
    접점 가격 검증
    
    Args:
        S19: 과거 19일 종가 합계
        p: 검증할 가격
    
    Returns:
        접점 조건 만족 여부
    """
    # MA20과 엔벨로프 계산
    ma20 = (S19 + p) / 20
    envelope = ma20 * 0.8  # -20%
    
    # 매수선 계산 (엔벨로프의 윗 호가)
    buy1 = ceil_tick(envelope)
    
    # 접점 조건: 현재가 = 매수선
    return p == buy1


# ==================== 단위 테스트 ====================
def test_contact_price_calculator():
    """접점 가격 계산기 단위 테스트"""
    
    print("=== 접점 가격 계산기 테스트 ===")
    
    # 테스트 케이스 1: 일반 케이스
    print("\n1. 일반 케이스")
    S19_1 = 1653800
    result_1 = solve_contact_price_debug(S19_1)
    print(f"S19 = {S19_1:,}")
    print(f"x_star = {result_1['x_star']:,.2f}")
    print(f"접점 가격 = {result_1['contact_price']:,}원")
    print(f"호가 단위 = {result_1['delta']}원")
    print(f"상한값 = {result_1['upper']:,.2f}")
    print(f"반복 횟수 = {result_1['iterations']}")
    
    # 검증
    is_valid = verify_contact_price(S19_1, result_1['contact_price'])
    print(f"검증 결과: {'PASS' if is_valid else 'FAIL'}")
    
    # 테스트 케이스 2: 경계 구간 테스트
    print("\n2. 경계 구간 테스트 (20,000원 근처)")
    S19_2 = 19995 * 24  # x_star = 19,995
    result_2 = solve_contact_price_debug(S19_2)
    print(f"S19 = {S19_2:,}")
    print(f"x_star = {result_2['x_star']:,.2f}")
    print(f"접점 가격 = {result_2['contact_price']:,}원")
    print(f"호가 단위 = {result_2['delta']}원")
    
    # 검증
    is_valid = verify_contact_price(S19_2, result_2['contact_price'])
    print(f"검증 결과: {'PASS' if is_valid else 'FAIL'}")
    
    # 테스트 케이스 3: 대형주 구간 테스트
    print("\n3. 대형주 구간 테스트 (500,000원 근처)")
    S19_3 = 500000 * 24  # x_star = 500,000
    result_3 = solve_contact_price_debug(S19_3)
    print(f"S19 = {S19_3:,}")
    print(f"x_star = {result_3['x_star']:,.2f}")
    print(f"접점 가격 = {result_3['contact_price']:,}원")
    print(f"호가 단위 = {result_3['delta']}원")
    
    # 검증
    is_valid = verify_contact_price(S19_3, result_3['contact_price'])
    print(f"검증 결과: {'PASS' if is_valid else 'FAIL'}")
    
    # 테스트 케이스 4: 실제 삼성전자 데이터 기반
    print("\n4. 실제 삼성전자 데이터 기반 테스트")
    # 삼성전자 20일선 87,505원에서 역산
    S19_samsung = 87505 * 20 - 96300  # 현재가 제외한 과거 19일 합계
    result_samsung = solve_contact_price_debug(S19_samsung)
    print(f"S19 = {S19_samsung:,}")
    print(f"x_star = {result_samsung['x_star']:,.2f}")
    print(f"접점 가격 = {result_samsung['contact_price']:,}원")
    print(f"호가 단위 = {result_samsung['delta']}원")
    
    # 검증
    is_valid = verify_contact_price(S19_samsung, result_samsung['contact_price'])
    print(f"검증 결과: {'PASS' if is_valid else 'FAIL'}")


if __name__ == "__main__":
    test_contact_price_calculator()
