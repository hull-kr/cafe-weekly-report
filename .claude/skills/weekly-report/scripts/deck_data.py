"""통합 엑셀에서 주간보고 장표용 숫자를 뽑아 JSON 으로 저장한다 (읽기 전용).

사용법:
  .venv\\Scripts\\python -X utf8 deck_data.py <이번주 통합.xlsx> <점검사항.json> <결과.json> [--prev <지난주 통합.xlsx>]

매출은 엑셀의 계산값에 기대지 않고 수량 × 단가로 직접 구한다 (재계산 전이어도 동작).
"""
import argparse
import json
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent
CFG = json.loads((HERE.parent / "references" / "config.json").read_text(encoding="utf-8"))
MAN = 10000
WEEKDAY = "월화수목금토일"


def man(won):
    # 만 원 단위 사사오입 (파이썬 round 는 0.5 를 짝수로 보내므로 쓰지 않는다)
    return int(won / MAN + 0.5) if won >= 0 else -int(-won / MAN + 0.5)


def load(path):
    df = pd.read_excel(path, sheet_name="통합_판매")
    df = df[df["메뉴"].notna() & (df["지점"] != "합계")].copy()
    df["날짜"] = pd.to_datetime(df["날짜"])
    df["매출"] = df["수량"] * df["단가"]
    return df


def rel(p):
    try:
        return Path(p).resolve().relative_to(Path.cwd().resolve()).as_posix()
    except ValueError:
        return Path(p).as_posix()


def pct(now, before):
    return None if not before else round((now - before) / before * 100, 1)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("merged")
    ap.add_argument("checks")
    ap.add_argument("out")
    ap.add_argument("--prev")
    a = ap.parse_args()

    df = load(a.merged)
    ck = json.loads(Path(a.checks).read_text(encoding="utf-8"))
    prev = load(a.prev) if a.prev else None
    total, qty = float(df["매출"].sum()), int(df["수량"].sum())
    s, e = pd.Timestamp(ck["시작"]), pd.Timestamp(ck["끝"])

    st = df.groupby("지점").agg(매출=("매출", "sum"), 영업일=("날짜", "nunique"))
    st["점유율"] = st["매출"] / total * 100
    st["일평균"] = st["매출"] / st["영업일"]
    st = st.sort_values("매출", ascending=False)

    mn = df.groupby("메뉴").agg(매출=("매출", "sum"), 수량=("수량", "sum"))
    mn["점유율"] = mn["매출"] / total * 100
    mn = mn.sort_values("매출", ascending=False)

    mix = df.pivot_table(index="지점", columns="메뉴", values="매출", aggfunc="sum").fillna(0)
    mix = mix.div(mix.sum(axis=1), axis=0) * 100
    spread = float((mix.max() - mix.min()).max()) if len(mix) > 1 else 0.0

    cat = {m["이름"]: m.get("분류", "") for m in CFG["메뉴"]}
    dessert_menus = [m for m in mn.index if cat.get(m) == "디저트"]
    dessert = float(mn.loc[dessert_menus, "점유율"].sum()) if dessert_menus else 0.0

    df["주말"] = df["날짜"].dt.day_of_week >= 5
    g = df.groupby(["지점", "주말"]).agg(매출=("매출", "sum"), 일수=("날짜", "nunique"))
    avg = (g["매출"] / g["일수"]).unstack()
    days = []
    for name in st.index:
        wd = avg.loc[name].get(False) if name in avg.index else None
        we = avg.loc[name].get(True) if name in avg.index else None
        if pd.isna(wd) or pd.isna(we) or not wd:
            continue  # 평일이나 주말 기록이 아예 없으면 배율을 낼 수 없다
        r = float(we / wd)
        days.append({"이름": name, "배율": round(r, 2),
                     "유형": "주말형" if r >= 1.2 else ("오피스형" if r < 0.9 else "고른형"),
                     "평일_만": man(wd), "주말_만": man(we)})
    days.sort(key=lambda x: -x["배율"])
    no_ratio = [n for n in st.index if n not in {d["이름"] for d in days}]

    wow = None
    if prev is not None:
        # 지점이 빠지거나 새로 생기면 총액끼리 비교하면 안 된다 - 양쪽 주에 다 있는 지점끼리만 비교한다
        prev_stores = set(prev["지점"].unique())
        common = [k for k in st.index if k in prev_stores]
        cur_c, prev_c = df[df["지점"].isin(common)], prev[prev["지점"].isin(common)]
        pst = prev.groupby("지점").agg(매출=("매출", "sum"), 영업일=("날짜", "nunique"))
        cmn, pmn = cur_c.groupby("메뉴")["매출"].sum(), prev_c.groupby("메뉴")["매출"].sum()
        stores = {}
        for k, r in st.iterrows():
            if k not in pst.index:
                stores[k] = {"증감": None, "일평균증감": None, "영업일다름": False, "전주영업일": None}
                continue
            p = pst.loc[k]
            # 영업일 수가 다르면 총액 증감은 휴무 탓이 섞인다 - 일평균 증감을 함께 낸다
            stores[k] = {"증감": pct(float(r.매출), float(p.매출)),
                         "일평균증감": pct(float(r.일평균), float(p.매출) / int(p.영업일)),
                         "영업일다름": int(r.영업일) != int(p.영업일), "전주영업일": int(p.영업일)}
        ps, pe = prev["날짜"].min(), prev["날짜"].max()
        wow = {
            "기간": f"{ps:%m.%d}~{pe:%m.%d}",
            "총매출_만": man(float(prev["매출"].sum())),
            "총매출_증감": pct(float(cur_c["매출"].sum()), float(prev_c["매출"].sum())),
            "총수량_증감": pct(int(cur_c["수량"].sum()), int(prev_c["수량"].sum())),
            "같은지점기준": set(common) != set(st.index) or set(common) != prev_stores,
            "빠진지점": sorted(prev_stores - set(st.index)),
            "신규지점": [k for k in st.index if k not in prev_stores],
            "지점": stores,
            "메뉴": {k: pct(float(cmn.get(k, 0)), float(pmn.get(k, 0))) for k in mn.index},
        }

    holidays = {k: [f"{pd.Timestamp(d):%m/%d}({WEEKDAY[pd.Timestamp(d).weekday()]})" for d in v]
                for k, v in ck.get("휴무", {}).items()}
    checks = [c for c in ck["점검"] if c["확인필요"]]
    fmt = [c for c in ck["점검"] if not c["확인필요"]]

    data = {
        "주": ck["주"],
        "기간": f"{s:%Y.%m.%d}({WEEKDAY[s.weekday()]}) ~ {e:%m.%d}({WEEKDAY[e.weekday()]})",
        "브랜드": CFG.get("브랜드", ""),
        "총매출_만": man(total), "총수량": qty,
        "지점": [{"이름": k, "매출_만": man(r.매출), "점유율": round(r.점유율, 1), "영업일": int(r.영업일),
                "일평균_만": man(r.일평균)} for k, r in st.iterrows()],
        "메뉴": [{"이름": k, "매출_만": man(r.매출), "수량": int(r.수량), "점유율": round(r.점유율, 1)}
               for k, r in mn.iterrows()],
        "메뉴구성_최대차이": round(spread, 1),
        "디저트": {"메뉴": dessert_menus, "점유율": round(dessert, 1)},
        "요일": days, "요일_제외": no_ratio,
        "휴무": holidays,
        "전주": wow,
        "점검": [{"지점": c["지점"], "유형": c["유형"], "내용": c["원본상태"], "처리": c["처리"],
                "영향_원": c["매출영향"], "관련금액_원": c.get("관련금액", 0)} for c in checks],
        "형식통일": {"건수": len(fmt), "지점": sorted({c["지점"] for c in fmt})},
        "메모": ck.get("메모", []),
        # 장표에 적을 엑셀 위치 - 실제로 만든 곳 (맑은 고딕은 역슬래시를 ₩ 로 그리므로 / 로 바꾼다)
        "엑셀경로": rel(a.merged),
    }
    Path(a.out).write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"성공: {a.out} (총매출 {data['총매출_만']:,}만 원, 점검 {len(checks)}건, 전주 대비 {'있음' if wow else '없음'})")


if __name__ == "__main__":
    main()
