// 주간 판매 보고 장표 생성 (숫자는 deck_data.py 가 만든 JSON 에서 읽는다)
// 사용법: node build_deck.js <장표데이터.json> <결과.pptx>   (프로젝트 폴더의 node_modules 를 쓴다)
//
// PowerPoint 는 한글을 글자 단위로 줄바꿈해 단어가 잘린다. 그래서 글상자에 넣는 문장은 모두
// wrapKo() 로 글자 폭을 계산해 단어 사이에서 미리 줄을 나누고, 제목은 fitOne() 으로 한 줄에 맞춘다.
// 차트 XML 에 한글 글꼴(<a:ea>)을 덧붙이면 PowerPoint 가 파일을 열지 못하므로 손대지 않는다.
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const [dataPath, outPath] = process.argv.slice(2);
const D = JSON.parse(fs.readFileSync(dataPath, "utf8"));

// 회사 색 진초록 계열 - 아주 진한 초록(배경·글자) · 진초록(주색) · 황금(강조) · 회색
const C = {
  dark: "0E2A1F", main: "1B5E3F", accent: "C99A2E", tint: "EDF4EF", card: "1A3D2E",
  gray: "5F6B65", lgray: "AEBAB3", line: "D8E2DC", white: "FFFFFF", cream: "D3E4D9",
  up: "1B5E3F", down: "B23A3A",
};
const FONT = "맑은 고딕";
const won = (n) => n.toLocaleString("ko-KR");
const sgn = (p) => (p === null || p === undefined ? "-" : `${p > 0 ? "+" : ""}${p.toFixed(1)}%`);

// ---------- 한국어 글자 폭·줄바꿈·조사 ----------
function charW(ch) {
  const c = ch.codePointAt(0);
  if ((c >= 0xac00 && c <= 0xd7a3) || (c >= 0x3130 && c <= 0x318f) || c >= 0x4e00) return 0.95;
  if (ch === " ") return 0.3;
  if (/[0-9]/.test(ch)) return 0.56;
  if (/[A-Z%]/.test(ch)) return 0.68;
  if (/[a-z]/.test(ch)) return 0.52;
  return 0.38; // 문장부호
}
const textW = (s, pt) => [...s].reduce((a, ch) => a + charW(ch), 0) * pt / 72;
function wrapKo(text, wIn, pt) {
  const max = wIn * 0.94;
  // 숫자와 단위('135만 원', '3 건')는 떼지 않는다 - 줄바꿈하지 않는 공백(U+00A0)으로 묶는다
  const NB = String.fromCharCode(0xa0);
  const glued = String(text).replace(/(\d[\d,.]*(?:만|천|억)?) (원|건|개|일|배)/g, `$1${NB}$2`)
    .replace(/ ~ /g, ` ~${NB}`); // 물결표가 줄 끝에 홀로 남지 않게 뒤 단어와 붙인다
  return glued.split("\n").map((para) => {
    const lines = [];
    let cur = "";
    for (const w of para.split(" ")) {
      const t = cur ? `${cur} ${w}` : w;
      if (cur && textW(t, pt) > max) { lines.push(cur); cur = w; } else cur = t;
    }
    lines.push(cur);
    return lines.join("\n");
  }).join("\n");
}
const nLines = (s) => String(s).split("\n").length;
function fitOne(text, wIn, pt, minPt) {
  let p = pt;
  while (p > minPt && textW(text, p) > wIn * 0.97) p -= 1;
  return p;
}
function hasBatchim(word) {
  const ch = String(word).trim().slice(-1);
  const c = ch.codePointAt(0);
  if (c >= 0xac00 && c <= 0xd7a3) return (c - 0xac00) % 28 !== 0;
  if (/[0-9]/.test(ch)) return "013678".includes(ch);
  return false;
}
const josa = (w, a, b) => `${w}${hasBatchim(w) ? a : b}`; // josa("강남","이","가") → 강남이

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const png = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + png.toString("base64");
}
async function badge(slide, Comp, x, y, d, bg, fg) {
  slide.addShape("ellipse", { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg } });
  const p = d * 0.25;
  slide.addImage({ data: await icon(Comp, fg), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
}
async function header(slide, Comp, section, title, dark = false) {
  await badge(slide, Comp, 0.5, 0.36, 0.42, dark ? C.accent : C.dark, dark ? C.dark : C.accent);
  slide.addText(section, { x: 1.04, y: 0.36, w: 5, h: 0.42, margin: 0, fontFace: FONT,
    fontSize: 13, bold: true, color: C.accent, valign: "middle" });
  slide.addText(title, { x: 0.5, y: 0.88, w: 9, h: 0.6, margin: 0, fontFace: FONT,
    fontSize: fitOne(title, 9, 22, 16), bold: true, color: dark ? C.white : C.dark, valign: "middle" });
}
function footnote(slide, text, dark = false) {
  slide.addText(text, { x: 0.5, y: 5.13, w: 9, h: 0.3, margin: 0, fontFace: FONT,
    fontSize: fitOne(text, 9, 9, 7), color: dark ? C.lgray : C.gray, valign: "middle" });
}
// 검토 제안 카드 - 글 줄 수에 맞춰 높이를 정하고, 실제 높이를 돌려준다
async function proposal(slide, x, y, w, text, maxH) {
  const body = wrapKo(text, w - 0.3, 11);
  const h = Math.min(maxH, 0.7 + nLines(body) * 0.2);
  slide.addShape("roundRect", { x, y, w, h, rectRadius: 0.08, fill: { color: C.tint }, line: { color: C.tint } });
  await badge(slide, fa.FaLightbulb, x + 0.15, y + 0.15, 0.34, C.accent, C.white);
  slide.addText("검토 제안", { x: x + 0.58, y: y + 0.15, w: w - 0.7, h: 0.34, margin: 0,
    fontFace: FONT, fontSize: 11, bold: true, color: C.main, valign: "middle" });
  slide.addText(body, { x: x + 0.15, y: y + 0.58, w: w - 0.3, h: h - 0.68, margin: 0,
    fontFace: FONT, fontSize: 11, color: C.dark, valign: "top", lineSpacingMultiple: 1.1 });
  return h;
}
function chartBase(extra) {
  return Object.assign({
    barDir: "bar", catAxisOrientation: "maxMin",
    catAxisLabelFontFace: FONT, catAxisLabelFontSize: 11, catAxisLabelColor: C.dark,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    catAxisLineShow: false, showValue: true, dataLabelFontFace: FONT,
    dataLabelFontSize: 11, dataLabelColor: C.dark, barGapWidthPct: 45,
    titleFontFace: FONT, titleFontSize: 11, titleColor: C.gray,
  }, extra);
}
// 6장 매출 영향 칸: 판단으로 바뀐 금액. 0 이어도 환불·추정처럼 알려야 할 금액은 따로 적는다
function impactText(c) {
  // 사사오입은 부호와 무관하게 (Math.round(-5.5) 는 -5 라서 절댓값으로 반올림한다)
  const r = (x) => Math.sign(x) * Math.round(Math.abs(x));
  const fmt = (v) => (Math.abs(v) < 10000 ? `${v > 0 ? "+" : ""}${won(r(v))}원`
    : `${v > 0 ? "+" : ""}${won(r(v / 10000))}만 원`);
  if (c.영향_원) return fmt(c.영향_원);
  if (c.유형 === "추정 매출" && c.관련금액_원) return `추정 ${fmt(c.관련금액_원).replace("+", "")}`;
  if (c.유형 === "환불") return "원본대로 차감";
  return "-";
}
const holidayNote = () => Object.entries(D.휴무 || {})
  .map(([s, ds]) => `${s} ${ds.join("·")} 휴무 추정`).join(", ");

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.author = D.브랜드 || "";
  pres.title = `주간 판매 보고 ${D.기간}`;
  const st = D.지점, mn = D.메뉴, dw = D.요일, W = D.전주;
  const top = st[0], topMenu = mn[0];
  const lowAvg = [...st].sort((a, b) => a.일평균_만 - b.일평균_만)[0];
  const topAvg = [...st].sort((a, b) => b.일평균_만 - a.일평균_만)[0];
  const nChecks = D.점검.length;

  // ---------- 1. 표지 ----------
  let s = pres.addSlide();
  s.background = { color: C.dark };
  await badge(s, fa.FaCoffee, 0.7, 0.75, 0.9, C.accent, C.dark);
  if (D.브랜드) s.addText(D.브랜드, { x: 0.7, y: 1.95, w: 6, h: 0.45, margin: 0, fontFace: FONT,
    fontSize: 18, bold: true, color: C.accent, charSpacing: 2 });
  s.addText("주간 판매 보고", { x: 0.7, y: 2.45, w: 8.5, h: 0.9, margin: 0, fontFace: FONT,
    fontSize: 42, bold: true, color: C.white });
  s.addText(`${D.기간}  ·  ${st.length}개 지점`, { x: 0.7, y: 3.45, w: 8.5, h: 0.4, margin: 0,
    fontFace: FONT, fontSize: 16, color: C.cream });
  s.addText(st.map((r) => r.이름).join(" · "), { x: 0.7, y: 4.65, w: 8.5, h: 0.35,
    margin: 0, fontFace: FONT, fontSize: 12, color: C.lgray });
  s.addNotes(`${D.기간} ${st.length}개 지점 주간 판매 보고입니다. 금액은 만 원 단위 반올림입니다.`);

  // ---------- 2. 한눈에 보기 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  await header(s, fa.FaChartBar, "한눈에 보기",
    `총매출 ${won(D.총매출_만)}만 원, ${josa(top.이름, "과", "와")} ${josa(topMenu.이름, "이", "가")} 이끈 한 주`);
  const same = W && W.같은지점기준 ? "*" : ""; // 설명은 각주에
  const kpis = [
    [`${won(D.총매출_만)}만 원`, W ? `주간 총매출\n전주 대비 ${sgn(W.총매출_증감)}${same}` : "주간 총매출"],
    [`${won(D.총수량)}개`, W ? `주간 총 판매 수량\n전주 대비 ${sgn(W.총수량_증감)}${same}` : "주간 총 판매 수량"],
    [top.이름, `매출 1위 지점\n${won(top.매출_만)}만 원`],
    [`${topMenu.점유율}%`, `1위 메뉴 비중\n${topMenu.이름}`],
  ];
  kpis.forEach(([big, label], i) => {
    const x = 0.5 + i * 2.333;
    s.addShape("roundRect", { x, y: 1.75, w: 2.0, h: 1.3, rectRadius: 0.08,
      fill: { color: C.tint }, line: { color: C.tint } });
    s.addText(big, { x: x + 0.15, y: 1.9, w: 1.7, h: 0.6, margin: 0, fontFace: FONT,
      fontSize: fitOne(big, 1.7, 24, 16), bold: true, color: i === 3 ? C.accent : C.dark, valign: "middle" });
    s.addText(wrapKo(label, 1.7, 10), { x: x + 0.15, y: 2.52, w: 1.7, h: 0.5, margin: 0, fontFace: FONT,
      fontSize: 10, color: C.gray, valign: "top" });
  });
  const wk = dw.filter((r) => r.유형 === "주말형"), of = dw.filter((r) => r.유형 === "오피스형");
  // 지점 비교는 일평균 증감으로 한다 - 휴무로 영업일이 다르면 총액 증감이 왜곡되기 때문
  const dDesc = W ? Object.entries(W.지점).filter(([, v]) => v.일평균증감 !== null)
    .map(([k, v]) => [k, v.일평균증감]).sort((a, b) => b[1] - a[1]) : [];
  const sameNote = W && W.같은지점기준
    ? `(같은 지점 기준${W.빠진지점.length ? `, ${W.빠진지점.join("·")} 자료 없음` : ""})` : "";
  const msgs = [
    W && dDesc.length
      ? `총매출 전주 대비 ${sgn(W.총매출_증감)}${sameNote} - 지점별 일평균은 ${dDesc[0][0]} ${sgn(dDesc[0][1])} ~ ${dDesc[dDesc.length - 1][0]} ${sgn(dDesc[dDesc.length - 1][1])}`
      : `${top.이름} 매출 1위(${top.점유율}%), 일평균 최저는 ${lowAvg.이름}(${lowAvg.일평균_만}만 원)`,
    `메뉴 1위 ${topMenu.이름}(${topMenu.점유율}%) - 지점 간 메뉴별 비중 차이 ${D.메뉴구성_최대차이}%p${D.메뉴구성_최대차이 < 5 ? " 이내로 거의 같음" : "로 지점마다 다름"}`,
    dw.length === 0 ? "주말·평일을 비교할 수 있는 지점이 없음"
      : wk.length && of.length
        ? `주말 매출은 지점 따라 정반대 - ${wk.map((r) => `${r.이름} ${r.배율}배`).join("·")}로 늘고, ${of.map((r) => `${r.이름} ${r.배율}배`).join("·")}로 줄어듦`
        : `주말/평일 매출 배율은 ${dw[0].이름} ${dw[0].배율}배 ~ ${dw[dw.length - 1].이름} ${dw[dw.length - 1].배율}배`,
  ];
  let y = 3.35;
  msgs.forEach((m, i) => {
    const t = wrapKo(m, 8.5, 14), h = 0.26 + nLines(t) * 0.24;
    s.addShape("ellipse", { x: 0.5, y: y + 0.04, w: 0.34, h: 0.34, fill: { color: C.dark }, line: { color: C.dark } });
    s.addText(String(i + 1), { x: 0.5, y: y + 0.04, w: 0.34, h: 0.34, margin: 0, fontFace: FONT,
      fontSize: 12, bold: true, color: C.accent, align: "center", valign: "middle" });
    s.addText(t, { x: 1.0, y, w: 8.5, h: Math.max(0.42, h - 0.08), margin: 0, fontFace: FONT, fontSize: 14,
      color: C.dark, valign: nLines(t) > 1 ? "top" : "middle" });
    y += Math.max(0.5, h);
  });
  footnote(s, [nChecks ? `사후 점검 사항 ${nChecks}건 - 마지막 장 참고` : "",
    W ? `${W.같은지점기준 ? "* " : ""}전주(${W.기간}) 대비${W.같은지점기준 ? ", 두 주 모두 자료가 있는 지점끼리 비교" : ""}` : "지난주 자료가 없어 전주 대비는 생략"]
    .filter(Boolean).join("  ·  "));
  s.addNotes(msgs.join(" "));

  // ---------- 3. 지점별 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  const low = st[st.length - 1];
  await header(s, fa.FaStore, "지점별", lowAvg.이름 === low.이름
    ? `${josa(top.이름, "이", "가")} 매출 1위, ${josa(low.이름, "은", "는")} 일평균으로 봐도 가장 낮음`
    : `${josa(top.이름, "이", "가")} 매출 1위, 일평균 최저는 ${lowAvg.이름}`);
  s.addChart(pres.charts.BAR, [{ name: "주간 매출(만 원)", labels: st.map((r) => r.이름),
    values: st.map((r) => r.매출_만) }], chartBase({ x: 0.4, y: 1.6, w: 5.4, h: 3.4,
    chartColors: [C.main], dataLabelFormatCode: '#,##0"만"', showTitle: true, title: "주간 매출 (만 원)",
    showLegend: false, valAxisMinVal: 0, valAxisMaxVal: Math.ceil(top.매출_만 / 100) * 100 + 100 }));
  // 머리글은 칸이 좁아도 한 줄에 들어가게 9pt (PowerPoint 는 '전주 대비' 를 글자 중간에서 끊는다)
  const th = { fontFace: FONT, fontSize: W ? 9 : 10, bold: true, color: C.white, fill: { color: C.dark },
    align: "center", valign: "middle", margin: 0.03 };
  const td = (t, extra = {}) => ({ text: t, options: Object.assign({ fontFace: FONT,
    fontSize: 11, color: C.dark, align: "center", valign: "middle" }, extra) });
  const head = ["지점", "점유율", "일평균"].concat(W ? ["전주 대비"] : []);
  const rowsT = [head.map((h) => ({ text: h, options: th }))];
  st.forEach((r) => {
    const row = [td(r.영업일 < 7 ? `${r.이름}*` : r.이름, { bold: true }), td(`${r.점유율.toFixed(1)}%`),
      td(`${r.일평균_만}만 원`)];
    if (W) {
      const v = W.지점[r.이름] || {};
      // 영업일 수가 전주와 다르면 총액 증감은 휴무 탓이 섞이므로 일평균 증감을 † 와 함께 보여 준다
      const shown = v.영업일다름 ? v.일평균증감 : v.증감;
      row.push(td(W.신규지점.includes(r.이름) ? "신규" : `${sgn(shown)}${v.영업일다름 ? "†" : ""}`,
        { color: shown > 0 ? C.up : shown < 0 ? C.down : C.dark, bold: true }));
    }
    rowsT.push(row);
  });
  const rowH = st.length > 6 ? 0.26 : 0.3;
  s.addTable(rowsT, { x: 6.1, y: 1.65, w: 3.4, colW: W ? [0.75, 0.8, 0.95, 0.9] : [1.0, 1.1, 1.3],
    rowH, border: { type: "solid", pt: 0.75, color: C.line } });
  const tableEnd = 1.65 + rowH * (st.length + 1);
  // 가장 많이 줄어든 지점은 일평균 증감으로 고른다 (휴무 하루 때문에 생긴 감소를 원인 확인 대상으로 오해하지 않게)
  const worst = dDesc.length ? dDesc[dDesc.length - 1] : null;
  const r3 = lowAvg.일평균_만 / topAvg.일평균_만;
  const prop3 = worst && worst[1] <= -10
    ? `${worst[0]} 일평균 매출이 전주 대비 ${sgn(worst[1])}로 가장 많이 줄어듦. 일시적 요인인지 원인 확인 검토`
    : `${josa(lowAvg.이름, "은", "는")} 일평균 ${lowAvg.일평균_만}만 원으로 ${topAvg.이름}(${topAvg.일평균_만}만 원)의 ${r3 > 0.45 && r3 < 0.55 ? "절반" : `${Math.round(r3 * 100)}%`} 수준. 상권·영업 시간 검토`;
  await proposal(s, 6.1, Math.max(3.65, tableEnd + 0.3), 3.4, prop3, 5.05 - Math.max(3.65, tableEnd + 0.3));
  const daysDiff = W ? Object.entries(W.지점).filter(([, v]) => v.영업일다름) : [];
  footnote(s, [st.some((r) => r.영업일 < 7) ? `* ${holidayNote() || "영업일 7일 미만"}` : "",
    daysDiff.length ? `† 영업일 수가 전주와 달라 일평균 기준 (총매출 기준 ${daysDiff.map(([k, v]) => `${k} ${sgn(v.증감)}`).join(", ")})` : "",
    "일평균 = 주간 매출 ÷ 영업일", W ? `전주 대비 = ${W.기간} 매출 대비` : "만 원 단위 반올림"]
    .filter(Boolean).join("  ·  "));
  s.addNotes(`${top.이름} ${won(top.매출_만)}만 원으로 1위입니다. ${prop3}`);

  // ---------- 4. 메뉴별 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  await header(s, fa.FaMugHot, "메뉴별", `${josa(topMenu.이름, "이", "가")} 매출의 ${topMenu.점유율}%, 메뉴 구성은 지점마다 ${D.메뉴구성_최대차이 < 5 ? "같음" : "다름"}`);
  s.addText(`${topMenu.점유율}%`, { x: 0.5, y: 1.6, w: 3.1, h: 0.7, margin: 0,
    fontFace: FONT, fontSize: 40, bold: true, color: C.accent, valign: "middle" });
  const l1 = `${topMenu.이름} 매출 비중 (${won(topMenu.매출_만)}만 원)`;
  s.addText(l1, { x: 0.5, y: 2.3, w: 3.1, h: 0.3, margin: 0, fontFace: FONT,
    fontSize: fitOne(l1, 3.1, 11, 9), color: C.gray });
  s.addText(`${D.메뉴구성_최대차이}%p`, { x: 0.5, y: 2.7, w: 3.1, h: 0.55, margin: 0,
    fontFace: FONT, fontSize: 30, bold: true, color: C.dark, valign: "middle" });
  s.addText("지점 간 메뉴별 매출 비중의 최대 차이", { x: 0.5, y: 3.25, w: 3.1, h: 0.3,
    margin: 0, fontFace: FONT, fontSize: 11, color: C.gray });
  const ds = D.디저트;
  await proposal(s, 0.5, 3.7, 3.1, ds.메뉴.length
    ? `디저트(${ds.메뉴.join("·")})는 매출의 ${ds.점유율}%. 커피와 묶은 세트 구성 검토`
    : `${topMenu.이름} 의존도가 ${topMenu.점유율}%. 다른 메뉴 판촉 검토`, 1.35);
  const menuH = Math.min(3.55, 0.5 + mn.length * 0.38);
  s.addChart(pres.charts.BAR, [{ name: "매출(만 원)", labels: mn.map((r) => r.이름),
    values: mn.map((r) => r.매출_만) }], chartBase({ x: 3.9, y: 1.55, w: 5.7, h: menuH,
    chartColors: mn.map((r, i) => (i === 0 ? C.accent : C.main)), dataLabelFormatCode: '#,##0"만"',
    showTitle: true, title: "메뉴별 주간 매출 (만 원)", showLegend: false,
    catAxisLabelFontSize: mn.length > 10 ? 8 : 10, valAxisMinVal: 0,
    valAxisMaxVal: Math.ceil(topMenu.매출_만 / 100) * 100 + 100 }));
  footnote(s, "만 원 단위 반올림  ·  메뉴 비중 차이 = 메뉴마다 지점별 매출 비중 중 최댓값 - 최솟값");
  s.addNotes(`${topMenu.이름} 하나가 매출의 ${topMenu.점유율}%입니다.`);

  // ---------- 5. 요일별 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  const last = dw[dw.length - 1];
  const frac = (r) => (r <= 0.4 ? "3분의 1" : r <= 0.6 ? "절반" : `${r}배`);
  await header(s, fa.FaCalendarAlt, "요일별", !dw.length ? "주말·평일 비교 불가"
    : wk.length && of.length
      ? `주말에 ${josa(wk.map((r) => r.이름).join("·"), "은", "는")} 늘고, ${josa(last.이름, "은", "는")} ${frac(last.배율)}로 줄어듦`
      : `주말/평일 매출 배율, ${dw[0].이름} ${dw[0].배율}배 ~ ${last.이름} ${last.배율}배`);
  const kinds = [["주말형", C.accent], ["고른형", C.lgray], ["오피스형", C.main]]
    .filter(([k]) => dw.some((r) => r.유형 === k));
  if (dw.length) {
    const maxR = Math.max(2, Math.ceil(dw[0].배율 * 2) / 2);
    s.addChart(pres.charts.BAR, kinds.map(([k]) => ({ name: k, labels: dw.map((r) => r.이름),
      values: dw.map((r) => (r.유형 === k ? r.배율 : 0)) })), chartBase({ x: 0.4, y: 1.55, w: 5.4, h: 3.5,
      barGrouping: "stacked", chartColors: kinds.map(([, c]) => c), dataLabelPosition: "inEnd",
      dataLabelColor: C.white, dataLabelFontBold: true, dataLabelFormatCode: '0.00"배";;',
      showTitle: true, title: "주말 일평균 ÷ 평일 일평균", showLegend: true, legendPos: "b",
      legendFontFace: FONT, legendFontSize: 10, legendColor: C.dark, valAxisHidden: false,
      valAxisMinVal: 0, valAxisMaxVal: maxR, valAxisMajorUnit: 0.5, valAxisLabelFormatCode: '0.0"배"',
      valAxisLabelFontFace: FONT, valAxisLabelFontSize: 9, valAxisLabelColor: C.gray,
      valGridLine: { color: C.line, size: 0.75 } }));
  }
  let gy = 1.6;
  for (const [k, col] of [["주말형", C.accent], ["오피스형", C.main], ["고른형", C.lgray]]) {
    const g = dw.filter((r) => r.유형 === k);
    if (!g.length) continue;
    s.addShape("ellipse", { x: 6.1, y: gy + 0.06, w: 0.2, h: 0.2, fill: { color: col }, line: { color: col } });
    s.addText(`${k}  ${g.map((r) => r.이름).join("·")}`, { x: 6.42, y: gy, w: 3.1, h: 0.32, margin: 0,
      fontFace: FONT, fontSize: 12, bold: true, color: C.dark, valign: "middle" });
    const t2 = `주말 ${g.map((r) => r.주말_만 + "만").join("·")} / 평일 ${g.map((r) => r.평일_만 + "만").join("·")} (일평균)`;
    s.addText(t2, { x: 6.42, y: gy + 0.32, w: 3.1, h: 0.28, margin: 0, fontFace: FONT,
      fontSize: fitOne(t2, 3.1, 10, 8), color: C.gray, valign: "middle" });
    gy += 0.68;
  }
  const parts = [];
  if (wk.length) parts.push(`${josa(wk.map((r) => r.이름).join("·"), "은", "는")} 주말 인력·재고 보강`);
  if (of.length) parts.push(`${josa(of[of.length - 1].이름, "은", "는")} 주말 영업 시간 조정`);
  const p5y = Math.max(3.7, gy + 0.1);
  await proposal(s, 6.1, p5y, 3.4, parts.length ? `${parts.join(", ")} 검토` : "요일별 차이가 크지 않아 현행 운영 유지", 5.05 - p5y);
  footnote(s, ["배율 = 주말(토·일) 일평균 매출 ÷ 평일(월~금) 일평균 매출", holidayNote() ? `휴무일은 빼고 계산 (${holidayNote()})` : "",
    D.요일_제외.length ? `${D.요일_제외.join("·")}: 평일 또는 주말 기록이 없어 제외` : ""].filter(Boolean).join("  ·  "));
  s.addNotes("주말 일평균을 평일 일평균으로 나눈 배율입니다. 유형 이름은 매출 흐름으로 붙인 것이고 상권 조사를 한 것은 아닙니다.");

  // ---------- 6. 사후 점검 사항 ----------
  s = pres.addSlide();
  s.background = { color: C.dark };
  await header(s, fa.FaClipboardCheck, "사후 점검 사항",
    nChecks ? `자동으로 판단해 처리한 이상 데이터 ${nChecks}건` : "판단이 필요한 이상 데이터 없음", true);
  // 주 폴더에 메모가 있으면 표 아래에 한 줄로 보여 준다 (판단 근거가 되는 경우가 많다)
  const memoText = (D.메모 || []).map((m) => m.내용.replace(/\s+/g, " ")).join(" / ");
  const tableMax = memoText ? 4.3 : 4.75;
  if (nChecks) {
    const cw = [0.9, 3.5, 3.4, 1.2];
    const hdr = ["지점", "원본 상태", "처리", "매출 영향"].map((h) => ({ text: h, options: {
      fontFace: FONT, fontSize: 10, bold: true, color: C.white, fill: { color: C.main }, align: "center", valign: "middle" } }));
    const rows = [hdr];
    let used = 1.62 + 0.3, shown = 0;
    for (const c of D.점검) {
      const a = wrapKo(c.내용, cw[1] - 0.15, 10), b = wrapKo(c.처리, cw[2] - 0.15, 10);
      const h = 0.12 + Math.max(nLines(a), nLines(b)) * 0.18;
      if (used + h > tableMax) break;
      const o = (extra = {}) => Object.assign({ fontFace: FONT, fontSize: 10, color: C.cream,
        fill: { color: C.card }, valign: "middle" }, extra);
      rows.push([{ text: c.지점, options: o({ align: "center", bold: true, color: C.white }) },
        { text: a, options: o() }, { text: b, options: o() },
        { text: impactText(c), options: o({ align: "center" }) }]);
      used += h; shown++;
    }
    s.addTable(rows, { x: 0.5, y: 1.62, w: 9, colW: cw, border: { type: "solid", pt: 0.75, color: C.dark } });
    if (shown < nChecks) {
      s.addText(`외 ${nChecks - shown}건 - 엑셀 '사후점검' 시트 참고`, { x: 0.5, y: Math.min(used + 0.05, tableMax + 0.03),
        w: 9, h: 0.25, margin: 0, fontFace: FONT, fontSize: 10, color: C.accent });
    }
  } else {
    s.addText("이번 주 자료에는 중복·환불·빠진 날 같은 판단이 필요한 데이터가 없었습니다.", { x: 0.5, y: 2.0, w: 9,
      h: 0.5, margin: 0, fontFace: FONT, fontSize: 14, color: C.cream });
  }
  if (memoText) {
    const mt = wrapKo(`폴더 메모: ${memoText}`, 9, 10).split("\n").slice(0, 2).join("\n");
    s.addText(mt, { x: 0.5, y: 4.38, w: 9, h: 0.42, margin: 0, fontFace: FONT, fontSize: 10,
      color: C.accent, valign: "top" });
  }
  await badge(s, fa.FaFileExcel, 0.5, 4.88, 0.32, C.main, C.cream);
  // 실제로 만든 엑셀 위치를 적는다. 맑은 고딕은 역슬래시를 ₩ 로 그리므로 경로는 / 로 적는다
  let xl = D.엑셀경로 || `보고서/${D.주}/통합_판매_${D.주}.xlsx`;
  if (textW(xl, 9) > 4.5) xl = xl.split('/').pop(); // 경로가 길면 파일 이름만 (각주가 두 줄로 넘어가지 않게)
  const ft = `전체 내역: ${xl} 의 '사후점검' 시트` +
    (D.형식통일.건수 ? `  ·  날짜·금액·메뉴 이름 등 형식만 바꾼 ${D.형식통일.건수}건 포함` : "");
  s.addText(ft, { x: 0.95, y: 4.88, w: 8.55, h: 0.32, margin: 0, fontFace: FONT,
    fontSize: fitOne(ft, 8.55, 10, 8), color: C.lgray, valign: "middle" });
  s.addNotes(nChecks ? `이상 데이터 ${nChecks}건은 묻지 않고 규칙대로 처리했습니다. 지점 확인이 필요하면 엑셀 사후점검 시트를 보시면 됩니다.`
    : "이번 주는 판단이 필요한 이상 데이터가 없었습니다.");

  await pres.writeFile({ fileName: outPath });
  console.log(`성공: ${outPath} 저장`);
}

main().catch((e) => { console.error("실패:", e); process.exit(1); });
