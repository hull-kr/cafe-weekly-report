# 원본 엑셀 구조 점검용 (읽기 전용)
import sys, glob, openpyxl
for f in sorted(glob.glob(sys.argv[1] + "/*.xlsx")):
    print("\n" + "="*20, f.split("/")[-1].split("\\")[-1], "="*20)
    wbf = openpyxl.load_workbook(f)
    wbv = openpyxl.load_workbook(f, data_only=True)
    for ws in wbf.worksheets:
        wv = wbv[ws.title]
        print(f"--- 시트 [{ws.title}] 상태={ws.sheet_state} 범위={ws.dimensions} 병합={list(map(str, ws.merged_cells.ranges))}")
        hid = [r for r, d in ws.row_dimensions.items() if d.hidden]
        hidc = [c for c, d in ws.column_dimensions.items() if d.hidden]
        if hid or hidc: print("  숨김 행:", hid, "숨김 열:", hidc)
        if ws.auto_filter.ref: print("  필터:", ws.auto_filter.ref)
        for row in ws.iter_rows():
            cells = []
            for c in row:
                if c.value is None: continue
                v = wv[c.coordinate].value
                s = f"{c.coordinate}={v!r}<{type(v).__name__}>"
                if isinstance(c.value, str) and c.value.startswith("="): s += f" 식[{c.value}]"
                if c.number_format != "General": s += f" 형식[{c.number_format}]"
                if c.comment: s += f" 메모[{c.comment.text.strip()}]"
                cells.append(s)
            if cells: print("  " + " | ".join(cells))
