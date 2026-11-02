# 다섯 지점 주간 판매 파일을 하나의 표(통합_판매.xlsx)로 합친다.
# 원본은 읽기만 한다. 사용법: .venv\Scripts\python -X utf8 tools\build_merged.py 1019-1025 통합_판매.xlsx
import re
import sys
from datetime import date

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SRC, OUT = sys.argv[1], sys.argv[2]
YEAR = 2026
MENU_ORDER = ["아이스 아메리카노", "따뜻한 아메리카노", "카페라떼", "바닐라라떼",
              "콜드브루", "딸기라떼", "크루아상", "치즈케이크"]
STORE_ORDER = ["강남", "대전", "성수", "판교", "해운대"]
ALIAS = {"아아": "아이스 아메리카노", "뜨아": "따뜻한 아메리카노", "라떼": "카페라떼",
         "바닐라": "바닐라라떼", "딸라": "딸기라떼", "치케": "치즈케이크"}
FILES = {
    "강남": "강남점 주간매출(10.19~10.25).xlsx",
    "대전": "대전점_주간_1019.xlsx",
    "성수": "성수점.xlsx",
    "판교": "판교_판매현황_1025.xlsx",
    "해운대": "해운대 매출 10월3주.xlsx",
}


def load(store):
    return openpyxl.load_workbook(f"{SRC}/{FILES[store]}").active


def md(text):
    m, d = map(int, re.match(r"\s*(\d+)/(\d+)", text).groups())
    return date(YEAR, m, d)


def won(text):
    return int(re.sub(r"[^\d-]", "", str(text)))


rows = []        # (지점, 날짜, 메뉴, 수량, 단가, 비고, 원본위치)
orig_total = {}  # 원본 그대로의 매출 합계 (수정 전)
orig_rows = {}

# 강남: 매출 열이 계산값 없는 수식이라 원본 합계도 수량×단가로 구한다
ws = load("강남")
orig_total["강남"], orig_rows["강남"] = 0, 0
for r in ws.iter_rows(min_row=2):
    if r[0].value is None:
        continue
    q, p = r[2].value, r[3].value
    orig_total["강남"] += q * p
    orig_rows["강남"] += 1
    rows.append(("강남", r[0].value.date(), r[1].value, q, p, "", f"{r[0].row}행"))

# 대전: 10/20 크루아상 중복 두 번째 줄 삭제, 10/23 카페라떼 -2 는 환불로 유지
ws = load("대전")
orig_total["대전"], orig_rows["대전"] = 0, 0
seen, removed = set(), []
for r in ws.iter_rows(min_row=2):
    if r[0].value is None:
        continue
    d, m, q, p, s = r[0].value.date(), r[1].value, r[2].value, r[3].value, r[4].value
    orig_total["대전"] += s
    orig_rows["대전"] += 1
    key = (d, m, q, p, s)
    if key in seen:
        removed.append(r[0].row)
        continue
    seen.add(key)
    note = "환불" if q < 0 else ""
    rows.append(("대전", d, m, q, p, note, f"{r[0].row}행"))
assert removed == [17], f"예상과 다른 중복 행: {removed}"

# 성수: 가로형(메뉴 × 날짜) → 세로형. 1행 제목, J열·11행 합계 제외
ws = load("성수")
hdr = [c.value for c in ws[2]]
orig_total["성수"], orig_rows["성수"] = 0, 0
for r in ws.iter_rows(min_row=3, max_row=10):
    menu, price = r[0].value, r[1].value
    for i in range(2, 9):
        q = r[i].value
        orig_total["성수"] += q * price
        orig_rows["성수"] += 1
        rows.append(("성수", md(hdr[i]), menu, q, price, "", r[i].coordinate))

# 판교: 1~2행 제목·작성자, 머리글 4행. 날짜는 '2026.10.19' 글자, 단가는 금액÷수량으로 역산
ws = load("판교")
orig_total["판교"], orig_rows["판교"] = 0, 0
for r in ws.iter_rows(min_row=5):
    if r[0].value is None:
        continue
    menu, dtxt, amount, q = (c.value for c in r[:4])
    price = amount // q
    assert price * q == amount, f"판교 {r[0].row}행 단가 역산 불일치"
    orig_total["판교"] += amount
    orig_rows["판교"] += 1
    y, m, d = map(int, dtxt.split("."))
    rows.append(("판교", date(y, m, d), menu, q, price, "", f"{r[0].row}행"))

# 해운대: 날짜 '10/19(월)' 글자(연도 없음), 금액 '4,500원' 글자, 메뉴 줄임말
ws = load("해운대")
orig_total["해운대"], orig_rows["해운대"] = 0, 0
for r in ws.iter_rows(min_row=2):
    if r[0].value is None:
        continue
    q, price, amount = r[2].value, won(r[3].value), won(r[4].value)
    assert price * q == amount, f"해운대 {r[0].row}행 금액 불일치"
    orig_total["해운대"] += amount
    orig_rows["해운대"] += 1
    rows.append(("해운대", md(r[0].value), ALIAS.get(r[1].value, r[1].value), q, price, "",
                 f"{r[0].row}행"))

# 메뉴명·단가 최종 점검
for st, d, m, q, p, *_ in rows:
    assert m in MENU_ORDER, f"모르는 메뉴: {st} {m}"
prices = {}
for _, _, m, _, p, *_ in rows:
    prices.setdefault(m, set()).add(p)
assert all(len(v) == 1 for v in prices.values()), f"지점 간 단가 불일치: {prices}"

# 정렬: 지점 → 날짜 → 메뉴 순서 → 판매 먼저, 환불 나중
rows.sort(key=lambda x: (STORE_ORDER.index(x[0]), x[1], MENU_ORDER.index(x[2]), x[3] < 0))

# ---------------- 엑셀 쓰기 ----------------
FONT = "맑은 고딕"
f_base = Font(name=FONT, size=10)
f_head = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_title = Font(name=FONT, size=12, bold=True)
f_bold = Font(name=FONT, size=10, bold=True)
f_red = Font(name=FONT, size=10, color="C00000")
fill_head = PatternFill("solid", fgColor="1F3864")
fill_sub = PatternFill("solid", fgColor="D9E1F2")
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center")
wrap = Alignment(vertical="top", wrap_text=True)

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "통합_판매"
heads = ["지점", "날짜", "요일", "메뉴", "수량", "단가", "매출", "비고", "원본 위치"]
widths = [9, 12, 6, 20, 8, 10, 12, 8, 11]
for i, (h, w) in enumerate(zip(heads, widths), 1):
    c = ws.cell(1, i, h)
    c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, box
    ws.column_dimensions[get_column_letter(i)].width = w

for n, (st, d, m, q, p, note, loc) in enumerate(rows, 2):
    vals = [st, d, f'=CHOOSE(WEEKDAY(B{n}),"일","월","화","수","목","금","토")',
            m, q, p, f"=E{n}*F{n}", note, loc]
    for i, v in enumerate(vals, 1):
        c = ws.cell(n, i, v)
        c.font, c.border = f_base, box
    ws.cell(n, 2).number_format = "yyyy-mm-dd"
    for col in (1, 2, 3, 8, 9):
        ws.cell(n, col).alignment = center
    for col in (5, 6, 7):
        ws.cell(n, col).number_format = "#,##0;-#,##0"
    if q < 0:
        for col in (5, 7, 8):
            ws.cell(n, col).font = f_red
last = len(rows) + 1
tot = last + 1
ws.cell(tot, 1, "합계").font = f_bold
ws.cell(tot, 5, f"=SUM(E2:E{last})")
ws.cell(tot, 7, f"=SUM(G2:G{last})")
for i in range(1, 10):
    c = ws.cell(tot, i)
    c.fill, c.border = fill_sub, box
    if i in (5, 7):
        c.font, c.number_format = f_bold, "#,##0;-#,##0"
ws.cell(tot, 1).alignment = center
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:I{last}"
ws["G1"].comment = Comment("매출 = 수량 × 단가 (모든 지점 공통으로 다시 계산)", "Claude")
ws["I1"].comment = Comment("각 지점 원본 파일에서 이 줄이 온 행 또는 셀", "Claude")

# ---------------- 수정내역 시트 ----------------
lg = wb.create_sheet("수정내역")
lg["A1"] = "통합 과정에서 고친 내용 (원본 폴더: 1019-1025, 원본 파일은 수정하지 않음)"
lg["A1"].font = f_title
log_head = ["번호", "지점", "원본 파일", "위치", "원본 상태", "처리", "근거"]
log_w = [6, 8, 30, 14, 42, 46, 26]
log = [
    ("공통", "전체", "-", "열 이름·순서가 지점마다 다름 (날짜/판매일자/일자, 수량/판매수량/개수, 매출/판매금액/합계금액 등)",
     "지점·날짜·요일·메뉴·수량·단가·매출·비고·원본 위치 로 통일", "합치기 위한 표준화"),
    ("공통", "전체", "-", "매출 값이 파일마다 수식·숫자·글자로 제각각",
     "모든 줄의 매출을 수식 =수량×단가 로 다시 계산", "단가가 전 지점 공통이고, 금액이 있던 지점(대전·판교·해운대)은 원본 금액과 전부 일치함을 확인"),
    ("강남", FILES["강남"], "E2:E57", "매출 열이 수식(=C×D)만 있고 계산된 값이 저장돼 있지 않음",
     "수량×단가 수식으로 다시 계산", "원본 값 그대로"),
    ("대전", FILES["대전"], "17행", "10/20 크루아상 15개·60,000원 줄이 16행과 똑같이 한 번 더 있음",
     "중복된 17행을 빼고 한 줄만 남김 (매출 -60,000원)", "사용자 지시 (2026-11-02)"),
    ("대전", FILES["대전"], "32행", "10/23 카페라떼 수량 -2, 매출 -10,000원",
     "환불로 보고 그대로 반영. 비고에 '환불' 표시, 빨간 글씨", "사용자 지시 (2026-11-02)"),
    ("대전", FILES["대전"], "-", "10/21(수) 판매 줄이 없음",
     "휴무일이라 그대로 둠 (빈 줄을 만들지 않음)", "사용자 확인: 그날 휴무 (2026-11-02)"),
    ("성수", FILES["성수"], "A1:J11", "가로형 표 (메뉴가 행, 날짜가 열). 1행 제목, J열·11행은 합계",
     "메뉴×날짜 한 칸을 한 줄로 풀어 세로형으로 바꿈. 제목·합계 행열은 제외", "합계는 통합 표에서 다시 구함"),
    ("성수", FILES["성수"], "C2:I2", "날짜가 '10/19(월)' 같은 글자이고 연도가 없음",
     f"{YEAR}년 날짜로 바꿈", f"파일 제목의 기간(10/19~10/25)과 요일이 {YEAR}년 달력과 맞음"),
    ("성수", FILES["성수"], "-", "매출 금액 없이 수량만 있음", "수량×단가로 매출 계산", "B열 단가 사용"),
    ("판교", FILES["판교"], "1~4행", "1행 제목, 2행 작성자, 머리글이 4행. 열 순서가 메뉴명·판매일자·판매금액·판매수량",
     "제목·작성자 줄을 빼고 표준 열 순서로 옮김", "-"),
    ("판교", FILES["판교"], "B5:B60", "날짜가 '2026.10.19' 같은 글자", "날짜 값으로 바꿈", "-"),
    ("판교", FILES["판교"], "C5:D60", "단가 열이 없음",
     "판매금액 ÷ 판매수량으로 단가를 구함", "56줄 모두 나누어떨어지고 다른 지점 단가와 같음"),
    ("해운대", FILES["해운대"], "A2:A57", "날짜가 '10/19(월)' 같은 글자이고 연도가 없음",
     f"{YEAR}년 날짜로 바꿈", f"파일명 '10월3주'와 요일이 {YEAR}년 달력과 맞음"),
    ("해운대", FILES["해운대"], "D2:E57", "가격·합계금액이 '4,500원' 같은 글자",
     "숫자로 바꿈", "56줄 모두 개수×가격 = 합계금액 확인"),
    ("해운대", FILES["해운대"], "B2:B57", "메뉴명이 줄임말",
     "정식 메뉴명으로 바꿈: " + ", ".join(f"{k}→{v}" for k, v in ALIAS.items()), "단가가 정식 메뉴와 같음"),
]
for i, (h, w) in enumerate(zip(log_head, log_w), 1):
    c = lg.cell(3, i, h)
    c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, box
    lg.column_dimensions[get_column_letter(i)].width = w
for n, item in enumerate(log, 4):
    for i, v in enumerate((n - 3, *item), 1):
        c = lg.cell(n, i, v)
        c.font, c.border, c.alignment = f_base, box, wrap
    lg.cell(n, 1).alignment = lg.cell(n, 2).alignment = center

# 원본 대비 합계 확인표
r0 = 4 + len(log) + 2
lg.cell(r0, 1, "원본 대비 합계 확인").font = f_title
chk_head = ["지점", "원본 줄 수", "통합 줄 수", "원본 매출 합계", "통합 매출 합계", "차이", "차이 설명"]
for i, h in enumerate(chk_head, 1):
    c = lg.cell(r0 + 1, i, h)
    c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, box
why = {"강남": "", "대전": "중복 크루아상 1줄(60,000원) 삭제", "성수": "",
       "판교": "", "해운대": ""}
for k, st in enumerate(STORE_ORDER):
    n = r0 + 2 + k
    vals = [st, orig_rows[st], f'=COUNTIF(통합_판매!$A$2:$A${last},A{n})',
            orig_total[st], f'=SUMIF(통합_판매!$A$2:$A${last},A{n},통합_판매!$G$2:$G${last})',
            f"=E{n}-D{n}", why[st]]
    for i, v in enumerate(vals, 1):
        c = lg.cell(n, i, v)
        c.font, c.border = f_base, box
        if i in (2, 3, 4, 5, 6):
            c.number_format = "#,##0;-#,##0;0"
    lg.cell(n, 1).alignment = center
nt = r0 + 2 + len(STORE_ORDER)
lg.cell(nt, 1, "합계")
for col in "BCDEF":
    lg[f"{col}{nt}"] = f"=SUM({col}{r0 + 2}:{col}{nt - 1})"
for i in range(1, 8):
    c = lg.cell(nt, i)
    c.font, c.fill, c.border = f_bold, fill_sub, box
    if 2 <= i <= 6:
        c.number_format = "#,##0;-#,##0;0"
lg.cell(nt, 1).alignment = center
note = lg.cell(nt + 1, 1,
               "※ '원본 줄 수·원본 매출 합계'는 원본 파일에서 읽어 적은 값입니다. "
               "강남·성수는 원본에 계산된 매출 값이 없어 수량×단가로 구했습니다.")
note.font = Font(name=FONT, size=9, color="595959")
lg.freeze_panes = "A4"

wb.save(OUT)
print(f"저장 완료: {OUT} (통합 {len(rows)}줄, 삭제한 대전 원본 행 {removed})")
