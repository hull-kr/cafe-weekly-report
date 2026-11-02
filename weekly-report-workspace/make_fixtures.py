# 스킬 시험용 가짜 주간 자료를 만든다 (지난주 실제 자료를 바탕으로 숫자를 바꾸고 이상 데이터를 심는다)
import random
import sys
from datetime import date, timedelta
from pathlib import Path

import openpyxl

random.seed(7)
ROOT = Path(sys.argv[1])
MENUS = [("아이스 아메리카노", 4500), ("따뜻한 아메리카노", 4500), ("카페라떼", 5000), ("바닐라라떼", 5500),
         ("콜드브루", 5000), ("딸기라떼", 6000), ("크루아상", 4000), ("치즈케이크", 6500)]
BASE = {"강남": 36, "대전": 18, "성수": 22, "판교": 34, "해운대": 19}  # 아이스 아메리카노 기준 평일 수량
WEEKEND = {"강남": 0.75, "대전": 1.0, "성수": 1.7, "판교": 0.38, "해운대": 1.55}
SHARE = [1.0, 0.47, 0.53, 0.37, 0.3, 0.16, 0.32, 0.19]


def qty(store, d, i, scale):
    w = WEEKEND[store] if d.weekday() >= 5 else 1.0
    return max(1, round(BASE[store] * 2.2 * SHARE[i] * w * scale * random.uniform(0.9, 1.1)))


def week(start, scale):
    return [start + timedelta(days=k) for k in range(7)]


def save(wb, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


# ---------------- 1026-1101 ----------------
W1 = ROOT / "1026-1101"
days = week(date(2026, 10, 26), 1)
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "판매"
ws.append(["날짜", "메뉴", "수량", "단가", "매출"])
for d in days:
    for i, (m, p) in enumerate(MENUS):
        name = "아이스아메리카노" if (i == 0 and d.weekday() in (1, 3)) else m
        ws.append([d, name, qty("강남", d, i, 1.04), p, None])
        r = ws.max_row; ws.cell(r, 5).value = f"=C{r}*D{r}"; ws.cell(r, 1).number_format = "yyyy-mm-dd"
save(wb, W1 / "강남점 주간매출(10.26~11.01).xlsx")

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "대전점"
ws.append(["날짜", "메뉴", "수량", "단가", "매출"])
ws.append([date(2026, 10, 25), "카페라떼", 24, 5000, 120000])  # 지난주 날짜가 섞임
tot_q = tot_a = 0
for d in days:
    for i, (m, p) in enumerate(MENUS):
        q = qty("대전", d, i, 0.97)
        ws.append([d, m, q, p, q * p]); tot_q += q; tot_a += q * p
        if d.day == 28 and m == "콜드브루":
            ws.append([d, m, q, p, q * p])  # 중복 줄
ws.append(["합계", None, tot_q, None, tot_a])
save(wb, W1 / "대전점_주간_1026.xlsx")

wb = openpyxl.Workbook(); ws = wb.active
ws.append(["성수점 10/26(월) ~ 11/01(일) 메뉴별 판매수량"])
ws.append(["메뉴", "단가"] + [f"{d.month}/{d.day}({'월화수목금토일'[d.weekday()]})" for d in days] + ["합계"])
for i, (m, p) in enumerate(MENUS + [("말차라떼", 6000)]):
    ii = min(i, 5)
    row = [m, p] + [qty("성수", d, ii, 1.02 if i < 8 else 0.4) for d in days]
    ws.append(row + [f"=SUM(C{ws.max_row + 1}:I{ws.max_row + 1})"])
save(wb, W1 / "성수점.xlsx")

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "판교점"
ws.append(["도담커피 판교점 판매현황 (10/26~11/01)"]); ws.append(["작성: 판교점 점장"]); ws.append([])
ws.append(["메뉴명", "판매일자", "판매금액", "판매수량"])
for d in days:
    for i, (m, p) in enumerate(MENUS):
        q = qty("판교", d, i, 0.95)
        amt = q * (5000 if (m == "바닐라라떼" and d.weekday() == 2) else p)  # 수요일 바닐라라떼 할인가처럼 보이는 금액
        ws.append([m, f"{d:%Y.%m.%d}", amt, q])
save(wb, W1 / "판교_판매현황_1101.xlsx")

ALIAS = {"아이스 아메리카노": "아아", "따뜻한 아메리카노": "뜨아", "카페라떼": "라떼", "바닐라라떼": "바닐라",
         "딸기라떼": "딸라", "치즈케이크": "치케"}
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "매출"
ws.append(["일자", "품목", "개수", "가격", "합계금액"])
for d in days[:6]:  # 일요일(11/01) 기록 없음
    for i, (m, p) in enumerate(MENUS):
        q = qty("해운대", d, i, 1.0)
        price_txt = "4,000원" if (m == "아이스 아메리카노" and d.weekday() == 4) else f"{p:,}원"  # 단가 칸 오타
        ws.append([f"{d.month}/{d.day}({'월화수목금토일'[d.weekday()]})", ALIAS.get(m, m), q, price_txt, f"{q * p:,}원"])
        if d.weekday() == 5 and m == "카페라떼":
            ws.append([f"{d.month}/{d.day}(토)", "라떼", -1, "5,000원", "-5,000원"])  # 환불
save(wb, W1 / "해운대 매출 10월4주.xlsx")
(W1 / "메모.txt").write_text("판교 수요일 바닐라라떼 행사 있었음", encoding="utf-8")

# ---------------- 1102-1108 ----------------
W2 = ROOT / "1102-1108"
days = week(date(2026, 11, 2), 1)
wb = openpyxl.Workbook(); ws = wb.active
ws.append(["Date", "Item", "Qty", "Unit Price", "Amount"])
for d in days:
    for i, (m, p) in enumerate(MENUS):
        q = qty("강남", d, i, 1.08); ws.append([d, m, q, p, q * p]); ws.cell(ws.max_row, 1).number_format = "yyyy-mm-dd"
save(wb, W2 / "gangnam_weekly_1102.xlsx")

wb = openpyxl.Workbook(); ws = wb.active
ws.append(["날짜", "메뉴", "수량", "단가", "매출"])
for d in days:
    for i, (m, p) in enumerate(MENUS):
        q = qty("대전", d, i, 1.0); ws.append([d, m, q, p, q * p])
save(wb, W2 / "대전점_주간_1102.xlsx")

wb = openpyxl.Workbook(); ws = wb.active
ws.append(["성수점 일자별 판매 (11/2~11/8)"])
for d in days:  # 날짜는 제목 줄로만, 그 아래 메뉴·수량 블록 - 스크립트가 못 읽는 형식
    ws.append([f"■ {d.month}월 {d.day}일 ({'월화수목금토일'[d.weekday()]})"])
    ws.append(["메뉴", "수량"])
    for i, (m, p) in enumerate(MENUS):
        ws.append([m, qty("성수", d, i, 0.98)])
    ws.append([])
save(wb, W2 / "성수점_1102.xlsx")

wb = openpyxl.Workbook(); ws = wb.active
ws.append(["메뉴명", "판매일자", "판매금액", "판매수량"])
for d in days:
    for i, (m, p) in enumerate(MENUS):
        q = qty("판교", d, i, 1.12); ws.append([m, f"{d:%Y-%m-%d}", q * p, q])
save(wb, W2 / "판교_판매현황_1108.xlsx")
print("시험 자료 생성:", W1, W2)
