# 시험 결과 채점 - 각 실행 폴더에 grading.json 을 쓴다 (skill-creator 뷰어 형식: text / passed / evidence)
import hashlib
import json
import sys
import zipfile
from pathlib import Path

import openpyxl

WS = Path(__file__).resolve().parent
ROOT = WS.parent
it = WS / (sys.argv[1] if len(sys.argv) > 1 else "iteration-1")


def check(text, ok, ev):
    return {"text": text, "passed": bool(ok), "evidence": str(ev)}


def common(out: Path, week: str):
    res = []
    xl = next(out.glob("통합_판매_*.xlsx"), None)
    pp = next(out.glob("주간보고_*.pptx"), None)
    res.append(check("통합 엑셀과 주간보고 PPT 가 모두 있음", xl and pp, f"{xl and xl.name}, {pp and pp.name}"))
    if not xl or not pp:
        return res, None, None
    wb = openpyxl.load_workbook(xl, data_only=True)
    res.append(check("통합 엑셀에 '통합_판매'·'사후점검' 시트가 있음", {"통합_판매", "사후점검"} <= set(wb.sheetnames), wb.sheetnames))
    g = [r[6] for r in wb["통합_판매"].iter_rows(min_row=2, values_only=True) if r[0] and r[0] != "합계"]
    res.append(check("매출 수식의 계산값이 모두 채워져 있음 (재계산 완료)", g and all(v is not None for v in g),
                     f"{sum(v is None for v in g)}칸 빈 값 / {len(g)}줄"))
    lg = wb["사후점검"]
    rows = list(lg.iter_rows(values_only=True))
    hi = next((i for i, r in enumerate(rows) if r and r[0] == "지점" and "설명 안 된 차이" in r), None)
    gaps = [r[7] for r in rows[hi + 1:] if hi is not None and r and r[0] and r[1] is not None and isinstance(r[7], (int, float))]
    res.append(check("원본 대비 검산의 '설명 안 된 차이'가 모두 0", hi is not None and gaps and all(abs(x) < 1 for x in gaps), gaps))
    z = zipfile.ZipFile(pp)
    slides = sorted(n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml"))
    res.append(check("PPT 가 6장", len(slides) == 6, len(slides)))
    pngs = list(out.rglob("slide-*.png"))
    res.append(check("실제 PowerPoint 로 열어 검수 그림 6장을 뽑음", len(pngs) >= 6, f"{len(pngs)}장"))
    last = z.read(f"ppt/slides/slide{len(slides)}.xml").decode("utf-8")
    res.append(check("마지막 장이 '사후 점검 사항'", "사후 점검 사항" in last, "마지막 장 텍스트 확인"))
    deck = json.loads((out / "장표데이터.json").read_text(encoding="utf-8")) if (out / "장표데이터.json").exists() else None
    ck = json.loads((out / "점검사항.json").read_text(encoding="utf-8")) if (out / "점검사항.json").exists() else None
    return res, deck, ck


def originals():
    bad = []
    for line in (WS / "원본해시.txt").read_text(encoding="utf-8").splitlines():
        h, p = line.split(maxsplit=1)
        p = p.lstrip("*")
        if hashlib.sha256((ROOT / p).read_bytes()).hexdigest() != h:
            bad.append(p)
    return check("원본·입력 파일이 바뀌지 않음", not bad, bad or "15개 모두 그대로")


def kinds(ck):
    return {(c["지점"], c["유형"]) for c in ck["점검"] if c["확인필요"]} if ck else set()


for run in sorted(it.glob("eval-*/with_skill")):
    out = run / "outputs"
    name = run.parent.name
    res, deck, ck = common(out, "")
    res.append(originals())
    report = (out / "보고.md").read_text(encoding="utf-8") if (out / "보고.md").exists() else ""
    k = kinds(ck)
    if "eval-1" in name:
        want = {("대전", "기간 밖 날짜"), ("대전", "중복 줄"), ("성수", "모르는 메뉴"), ("판교", "기준과 다른 단가"),
                ("해운대", "단가 오기"), ("해운대", "환불"), ("해운대", "빠진 날")}
        res.append(check("심어 둔 이상 데이터 7가지가 모두 사후 점검(확인 필요)에 있음", want <= k, sorted(want - k) or "7가지 모두"))
        memo = [c for c in (ck or {}).get("점검", []) if c["유형"] == "메모 파일"]
        res.append(check("메모.txt 를 확인 필요가 아닌 '메모 파일'로 기록", memo and not memo[0]["확인필요"] and ("공통", "읽지 않은 파일") not in k, memo[:1]))
        hw = (deck or {}).get("전주", {}).get("지점", {}).get("해운대", {})
        res.append(check("해운대 영업일 차이(휴무 추정)를 전주 대비에서 구분", isinstance(hw, dict) and hw.get("영업일다름"), hw))
        res.append(check("지난주(10.19~10.25) 대비가 장표에 들어감", deck and deck.get("전주") and deck["전주"]["기간"] == "10.19~10.25",
                         deck and deck.get("전주", {}) and deck["전주"].get("기간")))
        res.append(check("메모.txt 의 행사 내용을 보고에 반영", "행사" in report, "보고.md 에 '행사' 있음" if "행사" in report else "없음"))
    if "eval-2" in name:
        st = set(ck["지점"]) if ck else set()
        res.append(check("gangnam 파일을 강남으로 처리", "강남" in st and not any("gangnam" in s for s in st), sorted(st)))
        res.append(check("못 읽는 성수 블록형 파일을 변환해 포함", "성수" in st, sorted(st)))
        res.append(check("해운대 파일 없음을 사후 점검에 남김", ("해운대", "파일 없음") in k, sorted(k)))
        w = (deck or {}).get("전주") or {}
        res.append(check("빠진 지점 때문에 왜곡되지 않게 같은 지점끼리 전주 대비 (+6~8%)", w.get("같은지점기준") and w.get("총매출_증감") is not None and 6 <= w["총매출_증감"] <= 8, {"같은지점기준": w.get("같은지점기준"), "총매출_증감": w.get("총매출_증감")}))
        res.append(check("강남 이름 해석 판단을 사후 점검에 기록 (--note)", any(s == "강남" for s, _ in k), sorted(k)))
        res.append(check("지난주(10.26~11.01) 대비가 장표에 들어감", deck and deck.get("전주") and deck["전주"]["기간"] == "10.26~11.01",
                         deck and deck.get("전주", {}) and deck["전주"].get("기간")))
    if "eval-3" in name:
        res.append(check("총매출 3,511만 원 (손으로 만든 결과와 같음)", deck and deck["총매출_만"] == 3511, deck and deck["총매출_만"]))
        res.append(check("대전 판단 3건(중복·환불·빠진 날)만 확인 필요로 나옴",
                         k == {("대전", "중복 줄"), ("대전", "환불"), ("대전", "빠진 날")}, sorted(k)))
        res.append(check("지난주 자료가 없어 전주 대비 생략", deck and deck.get("전주") is None, deck and deck.get("전주")))
    passed = sum(r["passed"] for r in res)
    g = {"expectations": res, "summary": {"passed": passed, "failed": len(res) - passed, "total": len(res),
                                          "pass_rate": round(passed / len(res), 2)}}
    (run / "grading.json").write_text(json.dumps(g, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{name}: {passed}/{len(res)}")
    for r in res:
        if not r["passed"]:
            print(f"  실패: {r['text']} - {r['evidence']}")
