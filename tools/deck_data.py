# 통합_판매.xlsx 에서 주간보고 장표용 숫자를 뽑아 JSON 으로 저장한다 (읽기 전용)
# 사용법: .venv\Scripts\python -X utf8 tools\deck_data.py 통합_판매.xlsx tools\deck_data.json
import json
import sys

import pandas as pd

SRC, OUT = sys.argv[1], sys.argv[2]
df = pd.read_excel(SRC, sheet_name="통합_판매").dropna(subset=["메뉴"])
df["날짜"] = pd.to_datetime(df["날짜"])
assert df["매출"].notna().all(), "매출 계산값이 비어 있음 - recalc 먼저 실행"
MAN = 10000  # 만 원


def man(won):
    # 만 원 단위 사사오입 (파이썬 round 는 0.5 를 짝수로 보내므로 쓰지 않는다)
    return int(won / MAN + 0.5)

total = int(df["매출"].sum())
qty = int(df["수량"].sum())

st = df.groupby("지점").agg(매출=("매출", "sum"), 영업일=("날짜", "nunique"))
st["점유율"] = st["매출"] / total * 100
st["일평균"] = st["매출"] / st["영업일"]
st = st.sort_values("매출", ascending=False)

mn = df.groupby("메뉴").agg(매출=("매출", "sum"), 수량=("수량", "sum"))
mn["점유율"] = mn["매출"] / total * 100
mn = mn.sort_values("매출", ascending=False)

mix = df.pivot_table(index="지점", columns="메뉴", values="매출", aggfunc="sum")
mix = mix.div(mix.sum(axis=1), axis=0) * 100
mix_spread = float((mix.max() - mix.min()).max())

df["주말"] = df["날짜"].dt.day_of_week >= 5
g = df.groupby(["지점", "주말"]).agg(매출=("매출", "sum"), 일수=("날짜", "nunique"))
avg = (g["매출"] / g["일수"]).unstack()
ratio = (avg[True] / avg[False]).sort_values(ascending=False)


def kind(r):
    return "주말형" if r >= 1.2 else ("오피스형" if r < 0.9 else "고른형")


office = st.loc[["강남", "판교"], "매출"].sum() / total * 100
dessert = mn.loc[["크루아상", "치즈케이크"], "점유율"].sum()

data = {
    "기간": "2026.10.19(월) ~ 10.25(일)",
    "총매출_만": man(total),
    "총수량": qty,
    "지점": [
        {"이름": k, "매출_만": man(r.매출), "점유율": round(r.점유율, 1),
         "영업일": int(r.영업일), "일평균_만": man(r.일평균)}
        for k, r in st.iterrows()
    ],
    "메뉴": [
        {"이름": k, "매출_만": man(r.매출), "수량": int(r.수량), "점유율": round(r.점유율, 1)}
        for k, r in mn.iterrows()
    ],
    "메뉴구성_최대차이": round(mix_spread, 1),
    "디저트_점유율": round(dessert, 1),
    "오피스상권_점유율": round(office, 1),
    "요일": [
        {"이름": k, "배율": round(float(v), 2), "유형": kind(v),
         "평일_만": man(avg.loc[k, False]), "주말_만": man(avg.loc[k, True])}
        for k, v in ratio.items()
    ],
}
with open(OUT, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=1)
print(json.dumps(data, ensure_ascii=False, indent=1))
