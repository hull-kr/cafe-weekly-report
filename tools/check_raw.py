# 지점 파일 점검 (읽기 전용) - 공통 형태(지점, 날짜, 메뉴, 수량, 단가, 매출)로 임시 변환 후 검사
import re, sys, openpyxl, pandas as pd
D = sys.argv[1]
def num(s): return int(re.sub(r"[^\d-]", "", str(s))) if s is not None else None
rows = []
ws = openpyxl.load_workbook(f"{D}/강남점 주간매출(10.19~10.25).xlsx").active
for r in ws.iter_rows(min_row=2, values_only=True):
    rows.append(("강남", r[0].date(), r[1], r[2], r[3], None))  # 매출은 수식뿐
ws = openpyxl.load_workbook(f"{D}/대전점_주간_1019.xlsx").active
for r in ws.iter_rows(min_row=2, values_only=True):
    rows.append(("대전", r[0].date(), r[1], r[2], r[3], r[4]))
ws = openpyxl.load_workbook(f"{D}/성수점.xlsx").active
hdr = [c.value for c in ws[2]]
for r in ws.iter_rows(min_row=3, max_row=10, values_only=True):
    for i in range(2, 9):
        m, d = map(int, re.match(r"(\d+)/(\d+)", hdr[i]).groups())
        rows.append(("성수", pd.Timestamp(2026, m, d).date(), r[0], r[i], r[1], None))
ws = openpyxl.load_workbook(f"{D}/판교_판매현황_1025.xlsx").active
for r in ws.iter_rows(min_row=5, values_only=True):
    if r[0] is None: continue
    rows.append(("판교", pd.Timestamp(r[1].replace(".", "-")).date(), r[0], r[3], None, r[2]))
ws = openpyxl.load_workbook(f"{D}/해운대 매출 10월3주.xlsx").active
for r in ws.iter_rows(min_row=2, values_only=True):
    if r[0] is None: continue
    m, d = map(int, re.match(r"(\d+)/(\d+)", r[0]).groups())
    rows.append(("해운대", pd.Timestamp(2026, m, d).date(), r[1], r[2], num(r[3]), num(r[4])))
df = pd.DataFrame(rows, columns=["지점", "날짜", "메뉴", "수량", "단가", "매출"])
pd.set_option("display.width", 200); pd.set_option("display.max_columns", 20)
print("=== 지점·날짜별 행 수 ==="); print(df.pivot_table(index="날짜", columns="지점", values="수량", aggfunc="size"))
print("\n=== 지점별 메뉴명 ==="); 
for g, x in df.groupby("지점"): print(g, sorted(x["메뉴"].unique()))
print("\n=== 같은 날 같은 메뉴 중복 ==="); print(df[df.duplicated(["지점","날짜","메뉴"], keep=False)])
print("\n=== 수량 0 이하 ==="); print(df[df["수량"] <= 0])
print("\n=== 매출 != 수량*단가 ===")
c = df.dropna(subset=["단가","매출"]); print(c[c["매출"] != c["수량"]*c["단가"]])
p = df[df.지점=="판교"].assign(역산단가=lambda x: x.매출/x.수량); print("\n=== 판교 역산 단가 ==="); print(p.groupby("메뉴")["역산단가"].unique())
print("\n=== 메뉴별 단가(지점 간) ==="); print(df.dropna(subset=["단가"]).pivot_table(index="메뉴", columns="지점", values="단가", aggfunc=lambda s: sorted(set(s))))
print("\n=== 일별 총 수량 ==="); print(df.pivot_table(index="날짜", columns="지점", values="수량", aggfunc="sum"))
