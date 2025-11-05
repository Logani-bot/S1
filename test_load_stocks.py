#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
from Real_Time_Monitor_S1 import load_summary_stocks_with_buy_lines

print("=" * 80)
print("S1 실시간 모니터링 - 종목 로드 테스트")
print("=" * 80)

df = load_summary_stocks_with_buy_lines()

print(f"\n로드된 종목 수: {len(df)}")

if not df.empty:
    print("\n컬럼 목록:")
    print(list(df.columns))
    
    # SNT 다이내믹스 찾기
    snt = df[df['종목명'].str.contains('SNT|다이내믹스', na=False, case=False)]
    
    if not snt.empty:
        print(f"\n✅ SNT 다이내믹스 발견: {len(snt)}개")
        print("\nSNT 다이내믹스 정보:")
        for idx, row in snt.iterrows():
            print(f"  티커: {row['티커']}")
            print(f"  종목명: {row['종목명']}")
            print(f"  매수상태: {row['매수상태']}")
            print(f"  1차매수선: {row['1차매수선']:,.0f}")
            print(f"  2차매수선: {row['2차매수선']:,.0f}")
            print(f"  3차매수선: {row['3차매수선']:,.0f}")
    else:
        print("\n⚠ SNT 다이내믹스를 찾을 수 없습니다.")
        print("\n전체 종목 목록 (상위 10개):")
        print(df[['티커', '종목명', '1차매수선']].head(10).to_string())
else:
    print("\n❌ 종목을 로드할 수 없습니다.")

print("\n" + "=" * 80)


