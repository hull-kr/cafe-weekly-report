"""주 폴더(예: 1026-1101)의 지점별 판매 엑셀을 하나의 표로 합친다. 원본은 읽기만 한다.

사용법:
  .venv\\Scripts\\python -X utf8 merge_week.py <주폴더> <결과.xlsx> [--start 2026-10-26] [--adapter-dir 폴더]

- 머리글 이름(날짜/일자/판매일자, 메뉴/품목, 수량/개수 ...)과 표 모양(세로형/가로형)을 스스로 알아낸다.
- 이상 데이터는 묻지 않고 정해진 규칙으로 처리하고, 모든 판단을 '사후점검' 시트와 점검사항.json 에 남긴다.
- 읽지 못한 파일은 '실패'로 알리고 끝 코드 3 으로 끝난다. 그 파일은 Claude 가 표준 CSV
  (지점,날짜,메뉴,수량,단가,매출 / 날짜는 YYYY-MM-DD)로 바꿔 --adapter-dir 에 '<원본파일이름>.csv' 로 넣고 다시 돌린다.
"""
import argparse
import csv
import json
import re
import statistics
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
CFG = json.loads((HERE.parent / "references" / "config.json").read_text(encoding="utf-8"))
WEEKDAY = "월화수목금토일"


def norm(s) -> str:
    s = re.sub(r"\(.*?\)|\[.*?\]", "", str(s))  # 괄호 속 단위 표기 제거: 판매금액(원)
    return re.sub(r"[\s_\-·.:/]", "", s).lower()


FIELD_SYN = {
    "날짜": ["날짜", "일자", "판매일자", "판매일", "매출일자", "거래일자", "date", "day"],
    "메뉴": ["메뉴", "메뉴명", "품목", "품목명", "상품", "상품명", "제품", "제품명", "item", "menu", "product"],
    "수량": ["수량", "판매수량", "개수", "잔수", "판매량", "qty", "quantity", "count"],
    "단가": ["단가", "가격", "판매가", "판매단가", "price", "unitprice"],
    "매출": ["매출", "매출액", "판매금액", "합계금액", "금액", "판매액", "amount", "sales", "revenue"],
}
SYN2FIELD = {norm(w): f for f, ws in FIELD_SYN.items() for w in ws}
TOTAL_WORDS = {norm(w) for w in ["합계", "소계", "총계", "총합", "계", "total", "sum", "subtotal"]}
MENU = {m["이름"]: m for m in CFG["메뉴"]}
MENU_ALIAS = {norm(m["이름"]): m["이름"] for m in CFG["메뉴"]}
MENU_ALIAS.update({norm(a): m["이름"] for m in CFG["메뉴"] for a in m.get("별칭", [])})
STORES = [s["이름"] for s in CFG["지점"]]
STORE_ALIAS = sorted(((a, s["이름"]) for s in CFG["지점"] for a in [s["이름"], *s.get("별칭", [])]),
                     key=lambda x: -len(x[0]))


# ---------------------------------------------------------------- 기본 변환
def week_range(folder: str, start_arg: str | None):
    if start_arg:
        s = date.fromisoformat(start_arg)
        return s, s + timedelta(days=6)
    m = re.fullmatch(r"(\d{2})(\d{2})-(\d{2})(\d{2})", folder)
    if not m:
        sys.exit(f"실패: 폴더 이름 '{folder}' 에서 기간을 읽지 못함 (예: 1026-1101). --start 로 시작일을 주세요")
    sm, sd, em, ed = map(int, m.groups())
    today = date.today()
    cands = []
    for y in (today.year - 1, today.year, today.year + 1):
        s = date(y, sm, sd)
        e = date(y + (1 if (em, ed) < (sm, sd) else 0), em, ed)
        if s <= today + timedelta(days=7):
            cands.append((s, e))
    return max(cands)  # 오늘 기준으로 가장 최근의 해당 기간


def parse_date(v, start: date, end: date):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)) and 30000 < v < 80000:  # 엑셀 날짜 일련번호
        return date(1899, 12, 30) + timedelta(days=int(v))
    if v is None:
        return None
    s = str(v).strip()
    m = re.match(r"(\d{4})\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})", s)
    if m:
        try:
            return date(*map(int, m.groups()))
        except ValueError:
            return None
    m = re.match(r"(\d{1,2})\s*[.\-/월]\s*(\d{1,2})(?!\d)", s)
    if m:
        mo, d = map(int, m.groups())
        best = None
        for y in {start.year, end.year}:  # 연도가 없으면 이번 주 기간과 가장 가까운 해로 본다
            try:
                c = date(y, mo, d)
            except ValueError:
                continue
            gap = abs((c - start).days) if c < start else max(0, (c - end).days)
            if best is None or gap < best[0]:
                best = (gap, c)
        return best[1] if best else None
    return None


def to_num(v):
    if v is None or (isinstance(v, str) and not v.strip()):
        return None, False
    if isinstance(v, bool):
        return None, False
    if isinstance(v, (int, float)):
        return (int(v) if float(v).is_integer() else float(v)), False
    s = re.sub(r"[,\s원₩개잔]|krw", "", str(v), flags=re.I)
    s = s.replace("(", "-").replace(")", "") if s.startswith("(") else s
    if re.fullmatch(r"-?\d+(\.\d+)?", s):
        f = float(s)
        return (int(f) if f.is_integer() else f), True
    return None, False


def canon_menu(raw):
    return MENU_ALIAS.get(norm(raw))


def detect_store(path: Path, ws_list):
    for alias, name in STORE_ALIAS:
        if alias in path.stem:
            return name, False
    for ws in ws_list:  # 파일 이름에 없으면 첫 몇 줄의 제목에서 찾는다
        for row in ws.iter_rows(min_row=1, max_row=5, values_only=True):
            for v in row:
                if isinstance(v, str):
                    for alias, name in STORE_ALIAS:
                        if alias in v:
                            return name, False
    stem = re.sub(r"[\d_\-().~\s]|주간|매출|판매|현황|월|주|점$", "", path.stem) or path.stem
    return stem, True


# ---------------------------------------------------------------- 시트 읽기
def find_header(rows):
    """세로형: 날짜·메뉴·수량(또는 매출) 머리글이 있는 줄. 가로형: 메뉴 머리글 + 날짜 머리글 3개 이상."""
    for i, row in enumerate(rows[:15]):
        fields = {}
        for j, v in enumerate(row):
            if v is None:
                continue
            f = SYN2FIELD.get(norm(v))
            if f and f not in fields:
                fields[f] = j
        if "메뉴" in fields and "날짜" in fields and ("수량" in fields or "매출" in fields):
            return "세로형", i, fields
        if "메뉴" in fields:
            return "가로형?", i, fields
    return None, None, None


def read_sheet(ws, wsf, store, start, end, fname, notes):
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    frows = [list(r) for r in wsf.iter_rows(values_only=True)]
    kind, h, fields = find_header(rows)
    if kind is None:
        return None
    out = []

    def cell_note(r, c):
        fv = frows[r][c] if r < len(frows) and c < len(frows[r]) else None
        return isinstance(fv, str) and fv.startswith("=")

    if kind == "가로형?":
        date_cols = {}
        for j, v in enumerate(rows[h]):
            d = parse_date(v, start, end) if j != fields["메뉴"] and SYN2FIELD.get(norm(v or "")) is None else None
            if d:
                date_cols[j] = d
        if len(date_cols) < 3:
            return None
        kind = "가로형"
        mc, pc = fields["메뉴"], fields.get("단가")
        for i in range(h + 1, len(rows)):
            row = rows[i]
            menu = row[mc] if mc < len(row) else None
            if menu is None or not str(menu).strip():
                continue
            if norm(menu) in TOTAL_WORDS:
                notes["합계 줄 제외"].append(f"{i + 1}행")
                continue
            price, _ = to_num(row[pc]) if pc is not None and pc < len(row) else (None, False)
            for j, d in date_cols.items():
                q, txt = to_num(row[j] if j < len(row) else None)
                if q is None:
                    continue
                out.append(dict(store=store, date=d, menu_raw=str(menu).strip(), qty=q, price=price,
                                amount=None, file=fname,
                                loc=f"{get_column_letter(j + 1)}{i + 1}"))
        notes["형식"].append(f"가로형 표(메뉴×날짜), 머리글 {h + 1}행 - 세로형으로 풀어 씀")
        return out

    # 세로형
    last_date = None
    for i in range(h + 1, len(rows)):
        row = rows[i]
        g = lambda f: row[fields[f]] if f in fields and fields[f] < len(row) else None  # noqa: E731
        menu = g("메뉴")
        if menu is None or not str(menu).strip():
            continue
        if norm(menu) in TOTAL_WORDS or (isinstance(g("날짜"), str) and norm(g("날짜")) in TOTAL_WORDS):
            notes["합계 줄 제외"].append(f"{i + 1}행")
            continue
        dv = g("날짜")
        d = parse_date(dv, start, end)
        if d is None and dv in (None, "") and last_date:
            d = last_date  # 날짜가 첫 줄에만 적힌 표
            notes["빈 날짜 칸은 윗줄 날짜로 채움"].append(f"{i + 1}행")
        if isinstance(dv, str) and d:
            notes["날짜 글자를 날짜로 바꿈"].append(f"{i + 1}행")
        last_date = d or last_date
        q, qt = to_num(g("수량"))
        p, pt = to_num(g("단가"))
        a, at = to_num(g("매출"))
        if pt or at:
            notes["금액 글자를 숫자로 바꿈"].append(f"{i + 1}행")
        if "매출" in fields and a is None and cell_note(i, fields["매출"]):
            notes["매출 칸이 계산값 없는 수식 - 수량×단가로 다시 계산"].append(f"{i + 1}행")
        out.append(dict(store=store, date=d, date_raw=dv, menu_raw=str(menu).strip(), qty=q, price=p,
                        amount=a, file=fname, loc=f"{i + 1}행"))
    notes["형식"].append(f"세로형 표, 머리글 {h + 1}행 ({', '.join(sorted(fields))})")
    return out


def read_adapter_csv(path: Path, store, fname):
    out = []
    with open(path, encoding="utf-8-sig") as f:
        for n, r in enumerate(csv.DictReader(f), 2):
            out.append(dict(store=r.get("지점") or store, date=date.fromisoformat(r["날짜"]),
                            menu_raw=r["메뉴"].strip(), qty=to_num(r["수량"])[0],
                            price=to_num(r.get("단가"))[0], amount=to_num(r.get("매출"))[0],
                            file=fname, loc=f"변환CSV {n}행"))
    return out


# ---------------------------------------------------------------- 판단 규칙
class Checks:
    """사후 점검 기록. 같은 지점·유형·내용은 한 줄로 묶고 위치를 모은다."""

    def __init__(self):
        self.items = {}

    def add(self, store, file, kind, state, action, why, need_check, loc="", impact=0, key=None, rel=0):
        k = (store, kind, key or state)
        it = self.items.setdefault(k, dict(지점=store, 원본파일=file, 유형=kind, 원본상태=state, 처리=action,
                                           근거=why, 확인필요=need_check, 위치=[], 매출영향=0, 관련금액=0))
        if loc:
            it["위치"].append(loc)
        it["매출영향"] += impact
        it["관련금액"] += rel  # 추정 매출처럼 영향은 0 이지만 크기를 알려야 하는 금액

    def rows(self):
        def short(locs):
            return ", ".join(locs[:6]) + (f" 외 {len(locs) - 6}곳" if len(locs) > 6 else "")
        out = [dict(it, 위치=short(it["위치"])) for it in self.items.values()]
        out.sort(key=lambda x: (not x["확인필요"], STORES.index(x["지점"]) if x["지점"] in STORES else 99,
                                -abs(x["매출영향"])))
        return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("week_dir")
    ap.add_argument("out_xlsx")
    ap.add_argument("--start")
    ap.add_argument("--adapter-dir")
    ap.add_argument("--note", action="append",
                    help="직접 내린 판단 기록 '지점|유형|원본 상태|처리|근거|원본 파일' (여러 번 가능)")
    args = ap.parse_args()

    wdir = Path(args.week_dir)
    start, end = week_range(wdir.name, args.start)
    days = [start + timedelta(days=i) for i in range((end - start).days + 1)]
    ck = Checks()
    print(f"정보: 기간 {start} ~ {end} ({len(days)}일)")
    if start.weekday() != 0:
        ck.add("공통", "-", "기간", f"시작일 {start} 가 월요일이 아님", "폴더 이름의 기간을 그대로 사용",
               "폴더 이름 기준", True)

    files = sorted(p for p in wdir.iterdir() if p.suffix.lower() in (".xlsx", ".xlsm", ".csv")
                   and not p.name.startswith("~$"))
    others = [p for p in wdir.iterdir() if p.is_file() and p not in files and not p.name.startswith("~$")]
    memos = []
    for o in others:
        if o.suffix.lower() in (".txt", ".md"):
            text = o.read_text(encoding="utf-8", errors="replace").strip()
            memos.append({"파일": o.name, "내용": text})
            ck.add("공통", o.name, "메모 파일", text[:200], "판단 근거로 참고 (집계 대상 아님)", "주 폴더에 함께 들어온 메모", False)
        else:
            ck.add("공통", o.name, "읽지 않은 파일", f"{o.name} - 엑셀(.xlsx) 형식이 아님", "집계에서 제외",
                   ".xls 는 LibreOffice 로 .xlsx 변환 후 다시 돌리면 포함됨", True)
    for n in args.note or []:  # Claude 가 스크립트 밖에서 내린 판단 (별칭 추가 등)
        parts = (n.split("|") + [""] * 6)[:6]
        ck.add(parts[0] or "공통", parts[5] or "-", parts[1] or "직접 판단", parts[2], parts[3], parts[4] or "Claude 판단", True)

    raw, failed, seen_stores = [], [], defaultdict(list)
    for p in files:
        notes = defaultdict(list)
        adapter = Path(args.adapter_dir) / f"{p.name}.csv" if args.adapter_dir else None
        if p.suffix.lower() == ".csv" or (adapter and adapter.exists()):
            src = adapter if adapter and adapter.exists() else p
            store, newstore = detect_store(p, [])
            recs = read_adapter_csv(src, store, p.name)
            notes["형식"].append("Claude 가 표준 CSV 로 직접 변환한 자료 사용" if src != p else "CSV")
        else:
            wbv = openpyxl.load_workbook(p, data_only=True)
            wbf = openpyxl.load_workbook(p)
            sheets = [ws for ws in wbv.worksheets if ws.sheet_state == "visible"]
            store, newstore = detect_store(p, sheets)
            recs, used = [], []
            for ws in sheets:
                r = read_sheet(ws, wbf[ws.title], store, start, end, p.name, notes)
                if r:
                    recs += r
                    used.append(ws.title)
            hidden = [ws.title for ws in wbv.worksheets if ws.sheet_state != "visible"]
            if hidden:
                ck.add(store, p.name, "숨은 시트", f"숨은 시트 {hidden}", "읽지 않음", "보이는 시트만 집계", True)
            if len(used) > 1:
                ck.add(store, p.name, "여러 시트", f"자료가 있는 시트 {used}", "모두 합침",
                       "시트마다 다른 날짜·메뉴일 가능성이 큼 - 중복이면 아래 중복 규칙이 걸러냄", True)
        if not recs:
            failed.append(p.name)
            print(f"실패: {p.name} - 표 형식을 알아내지 못함")
            continue
        if newstore:
            ck.add(store, p.name, "새 지점", f"기준 지점 목록에 없는 이름 '{store}'", "파일 이름에서 지점 이름을 뽑아 집계",
                   "references/config.json 지점 목록에 추가 필요", True)
        seen_stores[store].append(p.name)
        for fm in notes["형식"]:
            if "가로형" in fm or "CSV" in fm:
                ck.add(store, p.name, "형식 통일", fm, "통합 표 형식으로 옮김", "값은 그대로", False)
        for k, locs in notes.items():
            if k == "형식":
                continue
            ck.add(store, p.name, "형식 통일", k, "자동 변환", "값은 그대로", False, loc="")
            ck.items[(store, "형식 통일", k)]["위치"] += locs
        print(f"정보: {p.name} → 지점 '{store}', {'; '.join(notes['형식'])}, {len(recs)}줄")
        raw += recs

    if failed:
        print("실패: 읽지 못한 파일 - 표준 CSV 로 변환해 --adapter-dir 에 넣고 다시 실행: " + ", ".join(failed))
        sys.exit(3)
    for s, fs in seen_stores.items():
        if len(fs) > 1:
            ck.add(s, ", ".join(fs), "같은 지점 파일 여러 개", f"{len(fs)}개 파일", "모두 합침 (완전히 같은 줄은 중복 규칙으로 제거)",
                   "추가 자료일 가능성", True)
    for s in STORES:
        if s not in seen_stores:
            ck.add(s, "-", "파일 없음", "이번 주 폴더에 이 지점 파일이 없음", "이번 주 집계에서 빠짐",
                   "지점에 자료 요청 필요", True)

    # 원본 그대로의 합계 (판단 전)
    orig = defaultdict(lambda: [0, 0])
    for r in raw:
        p0 = r["price"] or (MENU.get(canon_menu(r["menu_raw"]) or "", {}).get("단가"))
        r["raw_amt"] = r["amount"] if r["amount"] is not None else (r["qty"] or 0) * (p0 or 0)
        orig[r["store"]][0] += 1
        orig[r["store"]][1] += r["raw_amt"]

    kept = []
    for r in raw:
        st, f, loc = r["store"], r["file"], r["loc"]
        lost = -r["raw_amt"]
        if r["date"] is None:
            ck.add(st, f, "날짜 없음", f"날짜를 읽을 수 없는 줄 ({r.get('date_raw')!r})", "집계에서 제외",
                   "날짜를 알 수 없어 요일·일별 분석 불가", True, loc, lost)
            continue
        if not (start <= r["date"] <= end):
            ck.add(st, f, "기간 밖 날짜", f"이번 주 기간 밖 날짜 {r['date']}", "집계에서 제외",
                   "다른 주 자료가 섞인 것으로 판단", True, loc, lost, key=f"기간밖{r['date']}")
            continue
        menu = canon_menu(r["menu_raw"])
        if menu is None:
            menu = re.sub(r"\s+", " ", r["menu_raw"])
            ck.add(st, f, "모르는 메뉴", f"기준 메뉴에 없는 '{menu}'", "이름 그대로 새 메뉴로 집계",
                   "신메뉴·시즌 메뉴일 가능성. 맞으면 config.json 에 추가", True, loc, key=f"메뉴{menu}")
        elif menu != r["menu_raw"]:
            ck.add(st, f, "형식 통일", f"메뉴 이름 '{r['menu_raw']}'", f"'{menu}' 로 바꿈", "단가가 같음", False,
                   loc, key=f"별칭{r['menu_raw']}")
        q, p, a = r["qty"], r["price"], r["amount"]
        std = MENU.get(menu, {}).get("단가")
        if q is None:
            if a is not None and (p or std):
                q = a / (p or std)
                q = int(q) if float(q).is_integer() else q
                ck.add(st, f, "수량 없음", "수량 칸이 비어 있음", "매출 ÷ 단가로 수량을 구함", "금액은 있음", True, loc)
            else:
                ck.add(st, f, "수량 없음", "수량·매출이 모두 비어 있음", "집계에서 제외", "판매 기록으로 볼 수 없음",
                       True, loc, lost)
                continue
        if q == 0:
            ck.add(st, f, "형식 통일", "수량 0 인 줄", "제외", "매출 영향 없음", False, loc)
            continue
        derived = p is None and a is not None and q
        filep = p if p is not None else (abs(a / q) if derived else None)
        if filep is None and std is None:
            ck.add(st, f, "단가 없음", f"'{menu}' 단가·금액이 모두 없음", "집계에서 제외", "매출을 계산할 수 없음",
                   True, loc, lost)
            continue
        if filep is None:
            price = std
            ck.add(st, f, "추정 매출", "단가·금액 칸 없이 수량만 있음", "기준 단가로 매출 계산",
                   "할인·행사가 있었다면 실제 매출과 다름", True, loc, key="기준단가사용", rel=q * std)
        elif std is None or abs(filep - std) < 0.5:
            price = std if std is not None else filep
            if derived:
                ck.add(st, f, "형식 통일", "단가 칸 없음", "매출 ÷ 수량으로 단가를 구함", "기준 단가와 일치", False, loc,
                       key="단가역산")
            if a is not None and abs(a - q * price) > 0.5:
                ck.add(st, f, "금액 불일치", f"{menu} 매출 {a:,.0f}원 ≠ 수량 {q} × 단가 {price:,.0f}원",
                       "수량 × 단가로 다시 계산", "단가가 기준과 같아 금액 칸 입력 실수로 판단", True, loc,
                       q * price - a, key=f"금액{menu}{loc}")
        else:  # 파일의 단가가 기준 단가와 다름
            if a is not None and p is not None and abs(a - q * std) < 0.5:
                price = std
                ck.add(st, f, "단가 오기", f"{menu} 단가 {p:,.0f}원 (기준 {std:,.0f}원)", "기준 단가로 바로잡음",
                       "금액이 기준 단가와 맞아 단가 칸 입력 실수로 판단", True, loc, key=f"단가오기{menu}")
            else:
                price = filep
                ck.add(st, f, "기준과 다른 단가", f"{menu} {filep:,.0f}원 (기준 {std:,.0f}원)",
                       "파일의 단가(실제 판매가) 그대로 집계", "할인·가격 변경 가능성 - 지점 확인 필요", True, loc,
                       key=f"단가다름{menu}{filep}")
        note = ""
        if q < 0:
            note = "환불"
            ck.add(st, f, "환불", f"{r['date']:%m/%d} {menu} 수량 {q} ({q * price:,.0f}원)", "환불로 보고 그대로 차감 반영",
                   "지난 결정(2026-10 대전) 기준", True, loc, key=f"환불{loc}")
        kept.append(dict(store=st, date=r["date"], menu=menu, qty=q, price=price, note=note,
                         loc=f"{f} {loc}", file=f, raw_amt=r["raw_amt"]))

    # 같은 줄 중복
    final, seen = [], {}
    for r in kept:
        k = (r["store"], r["date"], r["menu"], r["qty"], r["price"])
        if k in seen and r["qty"] > 0:
            ck.add(r["store"], r["file"], "중복 줄", f"{r['date']:%m/%d} {r['menu']} {r['qty']}개가 똑같이 한 번 더 있음",
                   "한 줄만 남김", "복사 실수로 판단 (지난 결정 기준)", True, r["loc"].split(" ")[-1],
                   -r["raw_amt"], key=f"중복{k}")
            continue
        seen[k] = r
        final.append(r)
    # 판단으로 바뀐 금액(수량×단가와 원본 금액의 차이 포함)을 매출영향에 반영하기 위한 검산
    for r in final:
        r["amt"] = r["qty"] * r["price"]

    by_sdm = defaultdict(list)
    for r in final:
        if r["qty"] > 0:
            by_sdm[(r["store"], r["date"], r["menu"])].append(r)
    for (st, d, m), rs in by_sdm.items():
        if len(rs) > 1:
            ck.add(st, rs[0]["file"], "같은 메뉴 여러 줄", f"{d:%m/%d} {m} {len(rs)}줄 (수량 {', '.join(str(x['qty']) for x in rs)})",
                   "모두 합산", "수량이 달라 나눠 적은 것으로 판단", True)

    # 휴무 추정·일부 기록 누락·유난히 낮은 날
    holidays = {}
    for st in sorted({r["store"] for r in final}, key=lambda s: STORES.index(s) if s in STORES else 99):
        rs = [r for r in final if r["store"] == st]
        present = {r["date"] for r in rs}
        miss = [d for d in days if d not in present]
        if miss:
            holidays[st] = [d.isoformat() for d in miss]
            ck.add(st, rs[0]["file"], "빠진 날", "판매 기록이 없는 날 " + ", ".join(f"{d:%m/%d}({WEEKDAY[d.weekday()]})" for d in miss),
                   "휴무로 추정하고 영업일에서 뺌 (일평균은 영업일 기준)", "지난 결정(2026-10 대전 휴무) 기준", True)
        cnt = {d: len({r["menu"] for r in rs if r["date"] == d}) for d in present}
        mx = max(cnt.values())
        for d, c in sorted(cnt.items()):
            if c < mx * 0.6:
                ck.add(st, rs[0]["file"], "일부 기록만 있는 날", f"{d:%m/%d} 메뉴 {c}개만 기록 (평소 {mx}개)",
                       "있는 그대로 집계", "부분 영업 또는 입력 누락 가능성", True, key=f"부분{d}")
        daysum = defaultdict(float)
        for r in rs:
            daysum[r["date"]] += r["amt"]
        for wk in (False, True):
            grp = {d: v for d, v in daysum.items() if (d.weekday() >= 5) == wk}
            if len(grp) >= 2:
                med = statistics.median(grp.values())
                for d, v in grp.items():
                    if med > 0 and v < med * 0.5:
                        ck.add(st, rs[0]["file"], "유난히 낮은 날",
                               f"{d:%m/%d}({WEEKDAY[d.weekday()]}) 매출 {v / 1e4:,.0f}만 원 - 같은 {'주말' if wk else '평일'} 중앙값의 {v / med:.0%}",
                               "있는 그대로 집계", "단축 영업·입력 누락 가능성", True, key=f"낮음{d}")

    # 원본 대비 검산: 원본 합계 + 판단 영향 = 통합 합계 (판단 밖의 차이는 '금액 반올림/기타'로 드러남)
    merged_tot = defaultdict(float)
    for r in final:
        merged_tot[r["store"]] += r["amt"]
    impact = defaultdict(float)
    for it in ck.items.values():
        impact[it["지점"]] += it["매출영향"]
    for st in merged_tot:
        gap = merged_tot[st] - orig[st][1] - impact[st]
        if abs(gap) > 0.5:
            ck.add(st, seen_stores[st][0], "설명 안 된 차이", f"원본 합계와 통합 합계가 {gap:,.0f}원 다름",
                   "통합(수량×단가) 기준으로 집계", "원본 금액 칸과 수량×단가가 어긋난 줄이 있음", True, impact=gap)

    final.sort(key=lambda r: (STORES.index(r["store"]) if r["store"] in STORES else 99, r["date"],
                              list(MENU).index(r["menu"]) if r["menu"] in MENU else 99, r["qty"] < 0))
    write_xlsx(args.out_xlsx, final, ck.rows(), orig, start, end, wdir.name)

    checks = ck.rows()
    meta = dict(주=wdir.name, 시작=start.isoformat(), 끝=end.isoformat(), 휴무=holidays, 메모=memos,
                엑셀=str(Path(args.out_xlsx)),
                지점=sorted(merged_tot, key=lambda s: STORES.index(s) if s in STORES else 99),
                점검=checks, 확인필요_건수=sum(c["확인필요"] for c in checks),
                형식통일_건수=sum(not c["확인필요"] for c in checks))
    jpath = Path(args.out_xlsx).with_name("점검사항.json")
    jpath.write_text(json.dumps(meta, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    print(f"성공: {args.out_xlsx} ({len(final)}줄), 점검 {meta['확인필요_건수']}건 + 형식 통일 {meta['형식통일_건수']}건 → {jpath}")
    for c in checks:
        if c["확인필요"]:
            print(f"  점검: [{c['지점']}] {c['유형']} - {c['원본상태']} → {c['처리']}")


# ---------------------------------------------------------------- 엑셀 쓰기
def write_xlsx(path, rows, checks, orig, start, end, week):
    FONT = "맑은 고딕"
    f_base, f_bold = Font(name=FONT, size=10), Font(name=FONT, size=10, bold=True)
    f_head = Font(name=FONT, size=10, bold=True, color="FFFFFF")
    f_title, f_red = Font(name=FONT, size=12, bold=True), Font(name=FONT, size=10, color="C00000")
    fill_head, fill_sub = PatternFill("solid", fgColor="1B5E3F"), PatternFill("solid", fgColor="EDF4EF")
    fill_warn = PatternFill("solid", fgColor="FFF4D6")
    thin = Side(style="thin", color="BFBFBF")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    center, wrap = Alignment(horizontal="center", vertical="center"), Alignment(vertical="top", wrap_text=True)
    num = "#,##0;-#,##0;0"

    def head(ws, r, names, widths):
        for i, (h, w) in enumerate(zip(names, widths), 1):
            c = ws.cell(r, i, h)
            c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, box
            if w:
                ws.column_dimensions[get_column_letter(i)].width = w

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "통합_판매"
    head(ws, 1, ["지점", "날짜", "요일", "메뉴", "수량", "단가", "매출", "비고", "원본 위치"],
         [9, 12, 6, 20, 8, 10, 12, 8, 34])
    for n, r in enumerate(rows, 2):
        vals = [r["store"], r["date"], f'=CHOOSE(WEEKDAY(B{n}),"일","월","화","수","목","금","토")', r["menu"],
                r["qty"], r["price"], f"=E{n}*F{n}", r["note"], r["loc"]]
        for i, v in enumerate(vals, 1):
            c = ws.cell(n, i, v)
            c.font, c.border = f_base, box
        ws.cell(n, 2).number_format = "yyyy-mm-dd"
        for col in (1, 2, 3, 8):
            ws.cell(n, col).alignment = center
        for col in (5, 6, 7):
            ws.cell(n, col).number_format = num
        if r["qty"] < 0:
            for col in (5, 7, 8):
                ws.cell(n, col).font = f_red
    last = len(rows) + 1
    t = last + 1
    ws.cell(t, 1, "합계").alignment = center
    ws.cell(t, 5, f"=SUM(E2:E{last})")
    ws.cell(t, 7, f"=SUM(G2:G{last})")
    for i in range(1, 10):
        c = ws.cell(t, i)
        c.fill, c.border, c.font = fill_sub, box, f_bold
        if i in (5, 7):
            c.number_format = num
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:I{last}"
    ws["G1"].comment = Comment("매출 = 수량 × 단가 (모든 지점 공통으로 다시 계산)", "주간보고 스킬")

    lg = wb.create_sheet("사후점검")
    lg["A1"] = f"사후 점검 사항 - {start:%Y.%m.%d} ~ {end:%m.%d} (원본 폴더 {week}, 원본 파일은 수정하지 않음)"
    lg["A1"].font = f_title
    lg["A2"] = "노란 줄 = 자동으로 판단한 이상 데이터(확인 필요). 흰 줄 = 형식만 바꾼 것(값은 그대로)."
    lg["A2"].font = Font(name=FONT, size=9, color="595959")
    head(lg, 3, ["번호", "확인 필요", "지점", "원본 파일", "위치", "유형", "원본 상태", "처리", "근거", "매출 영향(원)"],
         [6, 8, 8, 26, 16, 14, 40, 34, 30, 12])
    for n, c in enumerate(checks, 4):
        vals = [n - 3, "예" if c["확인필요"] else "", c["지점"], c["원본파일"], c["위치"], c["유형"], c["원본상태"],
                c["처리"], c["근거"], c["매출영향"] or None]
        for i, v in enumerate(vals, 1):
            cell = lg.cell(n, i, v)
            cell.font, cell.border, cell.alignment = f_base, box, wrap
            if c["확인필요"]:
                cell.fill = fill_warn
        lg.cell(n, 10).number_format = num
        for col in (1, 2, 3):
            lg.cell(n, col).alignment = center
    r0 = 4 + len(checks) + 2
    lg.cell(r0, 1, "원본 대비 합계 확인").font = f_title
    head(lg, r0 + 1, ["지점", "원본 줄 수", "통합 줄 수", "원본 매출 합계", "통합 매출 합계", "차이", "판단으로 바뀐 금액",
                      "설명 안 된 차이"], [None] * 8)
    stores = sorted(orig, key=lambda s: STORES.index(s) if s in STORES else 99)
    for k, st in enumerate(stores):
        n = r0 + 2 + k
        vals = [st, orig[st][0], f"=COUNTIF(통합_판매!$A$2:$A${last},A{n})", orig[st][1],
                f"=SUMIF(통합_판매!$A$2:$A${last},A{n},통합_판매!$G$2:$G${last})", f"=E{n}-D{n}",
                f'=SUMIFS($J$4:$J${3 + len(checks)},$C$4:$C${3 + len(checks)},A{n},$F$4:$F${3 + len(checks)},"<>설명 안 된 차이")',
                f"=F{n}-G{n}"]
        for i, v in enumerate(vals, 1):
            cell = lg.cell(n, i, v)
            cell.font, cell.border = f_base, box
            if i > 1:
                cell.number_format = num
        lg.cell(n, 1).alignment = center
    note = lg.cell(r0 + 2 + len(stores), 1,
                   "※ '원본 매출 합계'는 원본 금액 칸(없으면 수량×단가)을 그대로 더한 값. '설명 안 된 차이'가 0 이 아니면 위 표의 해당 줄을 확인.")
    note.font = Font(name=FONT, size=9, color="595959")
    lg.freeze_panes = "A4"
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


if __name__ == "__main__":
    main()
