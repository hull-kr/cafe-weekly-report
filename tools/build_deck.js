// 주간 판매 보고 장표 생성 (숫자는 tools/deck_data.json 에서 읽는다)
// 사용법: node tools/build_deck.js tools/deck_data.json 주간보고_1019-1025.pptx
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const [dataPath, outPath] = process.argv.slice(2);
const D = JSON.parse(fs.readFileSync(dataPath, "utf8"));

// 색: 회사 색 진초록 계열 - 아주 진한 초록(배경·글자) · 진초록(주색) · 황금(강조) · 회색
const C = {
  dark: "0E2A1F", main: "1B5E3F", accent: "C99A2E", tint: "EDF4EF", card: "1A3D2E",
  gray: "5F6B65", lgray: "AEBAB3", line: "D8E2DC", white: "FFFFFF", cream: "D3E4D9",
};
const FONT = "맑은 고딕";
const won = (n) => n.toLocaleString("ko-KR");

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const png = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + png.toString("base64");
}

// 반복 요소: 원형 배지 안의 아이콘
async function badge(slide, Comp, x, y, d, bg, fg) {
  slide.addShape("ellipse", { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg } });
  const p = d * 0.25;
  slide.addImage({ data: await icon(Comp, fg), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
}

// 본문 장표 머리: 배지 + 구역 이름 + 핵심 문장 제목
async function header(slide, Comp, section, title) {
  await badge(slide, Comp, 0.5, 0.36, 0.42, C.dark, C.accent);
  slide.addText(section, { x: 1.04, y: 0.36, w: 4, h: 0.42, margin: 0, fontFace: FONT,
    fontSize: 13, bold: true, color: C.accent, valign: "middle" });
  slide.addText(title, { x: 0.5, y: 0.88, w: 9, h: 0.6, margin: 0, fontFace: FONT,
    fontSize: 22, bold: true, color: C.dark, valign: "middle" });
}

function footnote(slide, text) {
  slide.addText(text, { x: 0.5, y: 5.13, w: 9, h: 0.3, margin: 0, fontFace: FONT,
    fontSize: 9, color: C.gray, valign: "middle" });
}

// 검토 제안 카드
async function proposal(slide, x, y, w, h, text) {
  slide.addShape("roundRect", { x, y, w, h, rectRadius: 0.08, fill: { color: C.tint },
    line: { color: C.tint } });
  await badge(slide, fa.FaLightbulb, x + 0.15, y + 0.15, 0.34, C.accent, C.white);
  slide.addText("검토 제안", { x: x + 0.58, y: y + 0.15, w: w - 0.7, h: 0.34, margin: 0,
    fontFace: FONT, fontSize: 11, bold: true, color: C.main, valign: "middle" });
  slide.addText(text, { x: x + 0.15, y: y + 0.58, w: w - 0.3, h: h - 0.7, margin: 0,
    fontFace: FONT, fontSize: 11, color: C.dark, valign: "top", lineSpacingMultiple: 1.15 });
}

function chartBase(extra) {
  return Object.assign({
    barDir: "bar", catAxisOrientation: "maxMin",
    catAxisLabelFontFace: FONT, catAxisLabelFontSize: 11, catAxisLabelColor: C.dark,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    catAxisLineShow: false, showValue: true, dataLabelFontFace: FONT,
    dataLabelFontSize: 11, dataLabelColor: C.dark, barGapWidthPct: 45,
  }, extra);
}

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.author = "도담커피";
  pres.title = `주간 판매 보고 ${D.기간}`;
  const st = D.지점, mn = D.메뉴, dw = D.요일;
  const top = st[0], topMenu = mn[0];
  const byName = Object.fromEntries(dw.map((r) => [r.이름, r]));

  // ---------- 1. 표지 ----------
  let s = pres.addSlide();
  s.background = { color: C.dark };
  await badge(s, fa.FaCoffee, 0.7, 0.75, 0.9, C.accent, C.dark);
  s.addText("도담커피", { x: 0.7, y: 1.95, w: 6, h: 0.45, margin: 0, fontFace: FONT,
    fontSize: 18, bold: true, color: C.accent, charSpacing: 2 });
  s.addText("주간 판매 보고", { x: 0.7, y: 2.45, w: 8.5, h: 0.9, margin: 0, fontFace: FONT,
    fontSize: 42, bold: true, color: C.white });
  s.addText(`${D.기간}  ·  5개 지점`, { x: 0.7, y: 3.45, w: 8.5, h: 0.4, margin: 0,
    fontFace: FONT, fontSize: 16, color: C.cream });
  s.addText("강남 · 대전 · 성수 · 판교 · 해운대", { x: 0.7, y: 4.65, w: 8.5, h: 0.35,
    margin: 0, fontFace: FONT, fontSize: 12, color: C.lgray });
  s.addNotes(`${D.기간} 다섯 지점 주간 판매 보고입니다. 금액은 모두 만 원 단위 반올림입니다.`);

  // ---------- 2. 한눈에 보기 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  await header(s, fa.FaChartBar, "한눈에 보기",
    `총매출 ${won(D.총매출_만)}만 원, 강남과 아이스 아메리카노가 이끈 한 주`);
  const kpis = [
    [`${won(D.총매출_만)}만 원`, "주간 총매출"],
    [`${won(D.총수량)}개`, "주간 총 판매 수량"],
    [top.이름, `매출 1위 지점 · ${won(top.매출_만)}만 원`],
    [`${topMenu.점유율}%`, `1위 메뉴 비중
${topMenu.이름}`],
  ];
  const cw = 2.0, gap = 0.333;
  kpis.forEach(([big, label], i) => {
    const x = 0.5 + i * (cw + gap);
    s.addShape("roundRect", { x, y: 1.75, w: cw, h: 1.3, rectRadius: 0.08,
      fill: { color: C.tint }, line: { color: C.tint } });
    s.addText(big, { x: x + 0.15, y: 1.9, w: cw - 0.3, h: 0.6, margin: 0, fontFace: FONT,
      fontSize: 24, bold: true, color: i === 3 ? C.accent : C.dark, valign: "middle" });
    s.addText(label, { x: x + 0.15, y: 2.52, w: cw - 0.3, h: 0.42, margin: 0, fontFace: FONT,
      fontSize: 10, color: C.gray, valign: "top" });
  });
  const sungsu = byName["성수"], hwd = byName["해운대"], pg = byName["판교"];
  const msgs = [
    `주말에 매출이 줄어드는 강남·판교 두 곳이 전체 매출의 ${D.오피스상권_점유율}%`,
    `메뉴 구성은 다섯 지점이 거의 같음 - 메뉴별 매출 비중 차이 ${D.메뉴구성_최대차이}%p 이내`,
    `주말 매출은 지점 따라 정반대 - 성수 ${sungsu.배율}배·해운대 ${hwd.배율}배로 늘고, 판교는 ${pg.배율}배로 줄어듦`,
  ];
  msgs.forEach((m, i) => {
    const y = 3.4 + i * 0.52;
    s.addShape("ellipse", { x: 0.5, y: y + 0.04, w: 0.34, h: 0.34, fill: { color: C.dark },
      line: { color: C.dark } });
    s.addText(String(i + 1), { x: 0.5, y: y + 0.04, w: 0.34, h: 0.34, margin: 0, fontFace: FONT,
      fontSize: 12, bold: true, color: C.accent, align: "center", valign: "middle" });
    s.addText(m, { x: 1.0, y, w: 8.5, h: 0.42, margin: 0, fontFace: FONT, fontSize: 14,
      color: C.dark, valign: "middle" });
  });
  s.addNotes("핵심은 세 가지입니다. 지점 규모 차이, 지점마다 같은 메뉴 구성, 그리고 지점별로 정반대인 주말 흐름입니다.");

  // ---------- 3. 지점별 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  const dj = st.find((r) => r.이름 === "대전");
  await header(s, fa.FaStore, "지점별",
    `${top.이름}이 매출 1위, 대전은 일평균으로 봐도 가장 낮음`);
  s.addChart(pres.charts.BAR, [{ name: "주간 매출(만 원)",
    labels: st.map((r) => r.이름), values: st.map((r) => r.매출_만) }],
  chartBase({ x: 0.4, y: 1.6, w: 5.4, h: 3.4, chartColors: [C.main],
    dataLabelFormatCode: '#,##0"만"', showTitle: true, title: "주간 매출 (만 원)",
    titleFontFace: FONT, titleFontSize: 11, titleColor: C.gray, showLegend: false,
    valAxisMinVal: 0, valAxisMaxVal: Math.ceil(top.매출_만 / 100) * 100 + 100 }));
  const th = { fontFace: FONT, fontSize: 10, bold: true, color: C.white, fill: { color: C.dark },
    align: "center", valign: "middle" };
  const td = (t, extra = {}) => ({ text: t, options: Object.assign({ fontFace: FONT,
    fontSize: 11, color: C.dark, align: "center", valign: "middle" }, extra) });
  const rowsT = [[{ text: "지점", options: th }, { text: "점유율", options: th },
    { text: "일평균", options: th }]];
  st.forEach((r) => rowsT.push([
    td(r.영업일 < 7 ? `${r.이름}*` : r.이름, { bold: true }),
    td(`${r.점유율.toFixed(1)}%`), td(`${r.일평균_만}만 원`)]));
  s.addTable(rowsT, { x: 6.1, y: 1.65, w: 3.4, colW: [1.0, 1.1, 1.3], rowH: 0.3,
    border: { type: "solid", pt: 0.75, color: C.line } });
  await proposal(s, 6.1, 3.65, 3.4, 1.35,
    `대전은 휴무일을 빼도 일평균 ${dj.일평균_만}만 원으로\n${top.이름}(${top.일평균_만}만 원)의 절반 수준.\n상권·영업 시간 점검 검토`);
  footnote(s, "* 대전은 10/21(수) 휴무로 6일 영업  ·  일평균 = 주간 매출 ÷ 영업일  ·  만 원 단위 반올림");
  s.addNotes(`강남 ${won(top.매출_만)}만 원으로 1위입니다. 대전은 하루 휴무가 있어 일평균으로도 비교했는데, 일평균으로 봐도 가장 낮습니다.`);

  // ---------- 4. 메뉴별 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  await header(s, fa.FaMugHot, "메뉴별",
    `${topMenu.이름}가 매출의 ${topMenu.점유율}%, 메뉴 구성은 지점마다 같음`);
  s.addText(`${topMenu.점유율}%`, { x: 0.5, y: 1.6, w: 3.1, h: 0.7, margin: 0,
    fontFace: FONT, fontSize: 40, bold: true, color: C.accent, valign: "middle" });
  s.addText(`${topMenu.이름} 매출 비중 (${won(topMenu.매출_만)}만 원)`, { x: 0.5, y: 2.3,
    w: 3.1, h: 0.3, margin: 0, fontFace: FONT, fontSize: 11, color: C.gray });
  s.addText(`${D.메뉴구성_최대차이}%p`, { x: 0.5, y: 2.7, w: 3.1, h: 0.55, margin: 0,
    fontFace: FONT, fontSize: 30, bold: true, color: C.dark, valign: "middle" });
  s.addText("지점 간 메뉴별 매출 비중의 최대 차이", { x: 0.5, y: 3.25, w: 3.1, h: 0.3,
    margin: 0, fontFace: FONT, fontSize: 11, color: C.gray });
  // PowerPoint 는 한글을 글자 단위로 줄바꿈하므로 줄을 직접 나눈다 (한 줄이 카드 폭을 넘지 않게)
  await proposal(s, 0.5, 3.7, 3.1, 1.35,
    `디저트(크루아상·치즈케이크)는\n매출의 ${D.디저트_점유율}%.\n커피와 묶은 세트 구성 검토`);
  s.addChart(pres.charts.BAR, [{ name: "매출(만 원)", labels: mn.map((r) => r.이름),
    values: mn.map((r) => r.매출_만) }],
  chartBase({ x: 3.9, y: 1.55, w: 5.7, h: 3.55,
    chartColors: mn.map((r, i) => (i === 0 ? C.accent : C.main)),
    dataLabelFormatCode: '#,##0"만"', showTitle: true, title: "메뉴별 주간 매출 (만 원)",
    titleFontFace: FONT, titleFontSize: 11, titleColor: C.gray, showLegend: false,
    catAxisLabelFontSize: 10, valAxisMinVal: 0,
    valAxisMaxVal: Math.ceil(topMenu.매출_만 / 100) * 100 + 100 }));
  footnote(s, "만 원 단위 반올림  ·  메뉴 비중 차이 = 메뉴마다 다섯 지점의 매출 비중 중 최댓값 - 최솟값");
  s.addNotes("아이스 아메리카노 하나가 매출의 4분의 1이 넘습니다. 지점마다 메뉴 비중은 거의 같아서, 지점 간 차이는 메뉴가 아니라 판매량에서 납니다.");

  // ---------- 5. 요일별 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  await header(s, fa.FaCalendarAlt, "요일별",
    "주말에 성수·해운대는 늘고, 판교는 3분의 1로 줄어듦");
  const kinds = [["주말형", C.accent], ["고른형", C.lgray], ["오피스형", C.main]];
  s.addChart(pres.charts.BAR, kinds.map(([k]) => ({ name: k, labels: dw.map((r) => r.이름),
    values: dw.map((r) => (r.유형 === k ? r.배율 : 0)) })),
  chartBase({ x: 0.4, y: 1.55, w: 5.4, h: 3.5, barGrouping: "stacked",
    chartColors: kinds.map(([, c]) => c), dataLabelPosition: "inEnd", dataLabelColor: C.white,
    dataLabelFontBold: true, dataLabelFormatCode: '0.00"배";;', showTitle: true,
    title: "주말 일평균 ÷ 평일 일평균", titleFontFace: FONT, titleFontSize: 11,
    titleColor: C.gray, showLegend: true, legendPos: "b", legendFontFace: FONT,
    legendFontSize: 10, legendColor: C.dark, valAxisHidden: false, valAxisMinVal: 0,
    valAxisMaxVal: 2, valAxisMajorUnit: 0.5, valAxisLabelFormatCode: '0.0"배"',
    valAxisLabelFontFace: FONT, valAxisLabelFontSize: 9, valAxisLabelColor: C.gray,
    valGridLine: { color: C.line, size: 0.75 } }));
  const groups = [
    ["주말형", C.accent, ["성수", "해운대"]],
    ["오피스형", C.main, ["강남", "판교"]],
    ["고른형", C.lgray, ["대전"]],
  ];
  groups.forEach(([k, col, names], i) => {
    const y = 1.6 + i * 0.68;
    s.addShape("ellipse", { x: 6.1, y: y + 0.06, w: 0.2, h: 0.2, fill: { color: col },
      line: { color: col } });
    s.addText(`${k}  ${names.join("·")}`, { x: 6.42, y, w: 3.1, h: 0.32, margin: 0,
      fontFace: FONT, fontSize: 12, bold: true, color: C.dark, valign: "middle" });
    s.addText(`주말 ${names.map((n) => byName[n].주말_만 + "만").join("·")} / 평일 ${names.map((n) => byName[n].평일_만 + "만").join("·")} (일평균)`,
      { x: 6.42, y: y + 0.32, w: 3.1, h: 0.28, margin: 0, fontFace: FONT, fontSize: 10,
        color: C.gray, valign: "middle" });
  });
  await proposal(s, 6.1, 3.7, 3.4, 1.35,
    "성수·해운대는 주말 인력·재고 보강,\n판교는 주말 영업 시간 조정 검토");
  footnote(s, "배율 = 주말(토·일) 일평균 매출 ÷ 평일(월~금) 일평균 매출  ·  대전 평일은 휴무일(10/21) 빼고 4일 기준");
  s.addNotes("주말 일평균을 평일 일평균으로 나눈 배율입니다. 1보다 크면 주말에 더 팔리고, 작으면 주말에 덜 팔립니다. 유형 이름은 매출 흐름으로 붙인 것이고, 상권 조사를 한 것은 아닙니다.");

  // ---------- 6. 데이터 메모 ----------
  s = pres.addSlide();
  s.background = { color: C.dark };
  await badge(s, fa.FaClipboardCheck, 0.5, 0.36, 0.42, C.accent, C.dark);
  s.addText("데이터 메모", { x: 1.04, y: 0.36, w: 4, h: 0.42, margin: 0, fontFace: FONT,
    fontSize: 13, bold: true, color: C.accent, valign: "middle" });
  s.addText("이번 주 집계에 반영한 사항", { x: 0.5, y: 0.88, w: 9, h: 0.6, margin: 0,
    fontFace: FONT, fontSize: 22, bold: true, color: C.white, valign: "middle" });
  const memos = [
    [fa.FaCalendarTimes, "대전 10/21(수) 휴무", "6일 영업 기준으로 집계. 일평균도 6일로 나눔"],
    [fa.FaCopy, "대전 중복 1줄 삭제", "10/20 크루아상 15개(6만 원)가 두 번 입력돼\n한 줄만 반영"],
    [fa.FaUndoAlt, "대전 환불 1건 반영", "10/23 카페라떼 2개, 1만 원 차감"],
    [fa.FaTags, "지점별 형식 통일", "성수·해운대 날짜에 2026년 적용,\n해운대 메뉴 줄임말을 정식 이름으로"],
  ];
  memos.forEach(([Ic, t, d], i) => {
    const x = 0.5 + (i % 2) * 4.6, y = 1.75 + Math.floor(i / 2) * 1.25;
    s.addShape("roundRect", { x, y, w: 4.4, h: 1.05, rectRadius: 0.08,
      fill: { color: C.card }, line: { color: C.card } });
    // 동기 함수 안이라 배지는 아래에서 따로 그림
    s.addText(t, { x: x + 0.85, y: y + 0.15, w: 3.4, h: 0.35, margin: 0, fontFace: FONT,
      fontSize: 13, bold: true, color: C.white, valign: "middle" });
    s.addText(d, { x: x + 0.85, y: y + 0.5, w: 3.4, h: 0.45, margin: 0, fontFace: FONT,
      fontSize: 10, color: C.cream, valign: "top" });
  });
  for (let i = 0; i < memos.length; i++) {
    const x = 0.5 + (i % 2) * 4.6, y = 1.75 + Math.floor(i / 2) * 1.25;
    await badge(s, memos[i][0], x + 0.2, y + 0.25, 0.5, C.accent, C.dark);
  }
  await badge(s, fa.FaFileExcel, 0.5, 4.48, 0.36, C.main, C.cream);
  s.addText("원자료와 수정 내역: 통합_판매.xlsx 의 '수정내역' 시트  ·  다음 주부터 지난주 대비 증감 추가",
    { x: 1.0, y: 4.48, w: 8.5, h: 0.36, margin: 0, fontFace: FONT, fontSize: 11,
      color: C.lgray, valign: "middle" });
  s.addNotes("대전의 휴무, 중복 삭제, 환불은 지점 확인을 거쳐 반영했습니다. 자세한 내역은 통합 엑셀의 수정내역 시트에 있습니다.");

  // 주의: 차트 XML 에 한글 글꼴(<a:ea>)을 덧붙이면 PowerPoint 가 파일을 열지 못한다 (2026-11-02 확인).
  // 차트 숫자 '943 만' 처럼 벌어져 보이는 것은 LibreOffice 렌더링에서만 생기고 PowerPoint 에서는 붙어 나온다.
  await pres.writeFile({ fileName: outPath });
  console.log(`저장 완료: ${outPath}`);
}

main().catch((e) => { console.error("실패:", e); process.exit(1); });
