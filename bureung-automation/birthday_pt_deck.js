const pptxgen = require('pptxgenjs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';           // 13.33 x 7.5
pres.author = '피플실';
pres.title = '생일휴가 안내 업무 자동화';

// ===== Palette (VROONG brand green forward) =====
const DARK   = '0E2A22';   // deep forest — dark slides
const GREEN  = '00C878';   // VROONG accent
const WHITE  = 'FFFFFF';
const TINT   = 'F1F6F3';   // card background
const MUTED  = '5F7069';   // secondary text
const INK    = '15211D';   // body text on light
const WARN   = 'C7502F';   // As-Is problem marker

const F = 'Calibri';
const W = 13.33, H = 7.5;
const M = 0.75;                 // side margin
const CW = W - M * 2;           // content width

// ---- helpers (fresh objects every call) ----
function titleLight(s, txt, sub) {
  s.addText(txt, { x: M, y: 0.5, w: CW, h: 0.75, fontFace: F, fontSize: 38, bold: true,
                   color: INK, margin: 0, valign: 'middle' });
  if (sub) s.addText(sub, { x: M, y: 1.24, w: CW, h: 0.4, fontFace: F, fontSize: 14,
                            color: MUTED, margin: 0, valign: 'middle' });
}
function badge(s, x, y, label, d) {
  const dia = d || 0.5;
  s.addShape(pres.ShapeType.ellipse, { x: x, y: y, w: dia, h: dia, fill: { color: GREEN } });
  s.addText(label, { x: x, y: y, w: dia, h: dia, fontFace: F, fontSize: dia > 0.55 ? 16 : 14,
                     bold: true, color: WHITE, align: 'center', valign: 'middle', margin: 0 });
}
function card(s, x, y, w, h, fill) {
  s.addShape(pres.ShapeType.roundRect, { x: x, y: y, w: w, h: h, rectRadius: 0.08,
    fill: { color: fill || TINT }, line: { color: fill || TINT, width: 0 } });
}
function foot(s, txt) {
  s.addText(txt, { x: M, y: H - 0.5, w: CW, h: 0.3, fontFace: F, fontSize: 10,
                   color: MUTED, margin: 0, align: 'left' });
}

// =========================================================
// 1. TITLE (dark)
// =========================================================
let s = pres.addSlide();
s.background = { color: DARK };
s.addShape(pres.ShapeType.ellipse, { x: 10.7, y: -1.5, w: 4.6, h: 4.6, fill: { color: '173A2E' } });
s.addShape(pres.ShapeType.ellipse, { x: 11.9, y: 5.1, w: 2.6, h: 2.6, fill: { color: '173A2E' } });
s.addText('PEOPLE OPS AUTOMATION', { x: M, y: 2.05, w: 8.5, h: 0.35, fontFace: F, fontSize: 12,
  bold: true, color: GREEN, charSpacing: 3, margin: 0 });
s.addText('생일휴가 안내 업무 자동화', { x: M, y: 2.5, w: 9.6, h: 1.0, fontFace: F, fontSize: 46,
  bold: true, color: WHITE, margin: 0, valign: 'middle' });
s.addText('구글 스프레드시트 + Apps Script 기반 · 별도 비용 없음', {
  x: M, y: 3.55, w: 9.6, h: 0.45, fontFace: F, fontSize: 17, color: 'B8CFC4', margin: 0 });
s.addShape(pres.ShapeType.rect, { x: M, y: 4.35, w: 1.1, h: 0.045, fill: { color: GREEN } });
s.addText('피플실  |  2026. 08', { x: M, y: 4.65, w: 6, h: 0.35, fontFace: F, fontSize: 14,
  color: '8FA89B', margin: 0 });
s.addNotes('생일휴가 안내는 매월 반복되는 정형 업무입니다. 이번 발표는 이 업무를 자동화한 과정과, 담당자가 바뀌어도 계속 쓸 수 있게 만든 구조를 공유하는 자리입니다.');

// =========================================================
// 2. 목적
// =========================================================
s = pres.addSlide();
titleLight(s, '추진 목적', '반복 업무를 줄이고, 사람이 바뀌어도 유지되는 구조를 만든다');
const goals = [
  ['1', '반복 업무 자동화', '매월 동일하게 반복되는 명단 확인 · 대상자 선별 · 개별 메일 발송을 자동화'],
  ['2', '누락과 실수 제거', '규칙 기반으로 대상자를 선별하고 중복 발송을 시스템이 차단'],
  ['3', '속인화 해소', '담당자가 바뀌어도 같은 품질로 수행되도록 문서와 설정을 도구 안에 내장']
];
goals.forEach(function (g, i) {
  const y = 1.95 + i * 1.42;
  card(s, M, y, CW, 1.18);
  badge(s, M + 0.35, y + 0.34, g[0]);
  s.addText(g[1], { x: M + 1.05, y: y + 0.2, w: 3.4, h: 0.42, fontFace: F, fontSize: 19,
    bold: true, color: INK, margin: 0, valign: 'middle' });
  s.addText(g[2], { x: M + 4.55, y: y + 0.2, w: CW - 4.9, h: 0.78, fontFace: F, fontSize: 14,
    color: MUTED, margin: 0, valign: 'middle' });
});
foot(s, '대상 규모: 전 직원 180명 · 연간 안내 대상 161명 · 월평균 13.4명');
s.addNotes('목적은 세 가지입니다. 시간을 줄이고, 실수를 없애고, 사람이 바뀌어도 굴러가게 하는 것. 특히 세 번째가 이번 작업의 핵심입니다.');

// =========================================================
// 3. As-Is 문제
// =========================================================
s = pres.addSlide();
titleLight(s, '기존 방식의 문제', '매월 같은 작업을 손으로 반복해야 했습니다');
// process strip
const steps = ['ERP 명단 확인', '생일자 추출', '제외 대상 판별', '이메일 개별 조회', '메일 1건씩 작성', '발송 · 기록'];
const sw = (CW - 0.3 * 5) / 6;
steps.forEach(function (t, i) {
  const x = M + i * (sw + 0.3);
  card(s, x, 1.95, sw, 0.72, 'E9EFEB');
  s.addText(t, { x: x + 0.08, y: 1.95, w: sw - 0.16, h: 0.72, fontFace: F, fontSize: 12,
    color: INK, align: 'center', valign: 'middle', margin: 0 });
  if (i < 5) s.addText('>', { x: x + sw + 0.02, y: 1.95, w: 0.26, h: 0.72, fontFace: F,
    fontSize: 14, bold: true, color: MUTED, align: 'center', valign: 'middle', margin: 0 });
});
s.addText('모든 단계가 수작업 — 매월 반복', { x: M, y: 2.78, w: CW, h: 0.3, fontFace: F,
  fontSize: 12, italic: true, color: MUTED, margin: 0 });

const probs = [
  ['시간', '월 1.5~2시간', '13~22명에게 이메일 주소를 하나씩 확인하고 메일을 개별 작성'],
  ['정확성', '누락 · 오발송 위험', '휴직 · 퇴직 · 임원 제외를 매번 눈으로 대조'],
  ['속인화', '담당자 의존', '절차가 문서로 남지 않아 인수인계 시 재현이 어려움']
];
const pw = (CW - 0.4 * 2) / 3;
probs.forEach(function (p, i) {
  const x = M + i * (pw + 0.4);
  card(s, x, 3.35, pw, 2.45);
  s.addShape(pres.ShapeType.ellipse, { x: x + 0.32, y: 3.65, w: 0.42, h: 0.42, fill: { color: WARN } });
  s.addText('!', { x: x + 0.32, y: 3.65, w: 0.42, h: 0.42, fontFace: F, fontSize: 15, bold: true,
    color: WHITE, align: 'center', valign: 'middle', margin: 0 });
  s.addText(p[0], { x: x + 0.32, y: 4.2, w: pw - 0.64, h: 0.3, fontFace: F, fontSize: 12,
    bold: true, color: WARN, charSpacing: 1, margin: 0 });
  s.addText(p[1], { x: x + 0.32, y: 4.52, w: pw - 0.64, h: 0.45, fontFace: F, fontSize: 20,
    bold: true, color: INK, margin: 0, valign: 'middle' });
  s.addText(p[2], { x: x + 0.32, y: 5.02, w: pw - 0.64, h: 0.68, fontFace: F, fontSize: 13,
    color: MUTED, margin: 0 });
});
foot(s, '※ 이메일 주소가 사원명부에서 숨김 처리되어 사원등록 화면에서 1명씩 확인해야 했음');
s.addNotes('가장 큰 병목은 이메일 주소였습니다. 사원명부에서 이메일 열이 숨겨져 있어서, 한 명씩 사원등록 페이지에 들어가 확인해야 했습니다.');

// =========================================================
// 4. To-Be 개요
// =========================================================
s = pres.addSlide();
titleLight(s, '개선 방식', '스프레드시트 메뉴 4단계로 완료됩니다');
const flow = [
  ['0', 'ERP 명단 대조', 'ERP 최신본과 자동 비교\n휴직 · 퇴직 반영\n신규 입사자 자동 추가'],
  ['1', '대상 자동 분류', '생일월 자동 계산\n임원 · 휴직 · 제외 부서 자동 판별'],
  ['2', '테스트 발송', '본인 메일로 견본 발송\n문구 · 서명 사전 확인'],
  ['3', '일괄 발송', '개인화된 안내 메일 발송\n결과 자동 기록']
];
const fw = (CW - 0.35 * 3) / 4;
flow.forEach(function (f, i) {
  const x = M + i * (fw + 0.35);
  card(s, x, 2.0, fw, 2.75);
  badge(s, x + 0.3, 2.3, f[0], 0.6);
  s.addText(f[1], { x: x + 0.3, y: 3.02, w: fw - 0.6, h: 0.42, fontFace: F, fontSize: 17,
    bold: true, color: INK, margin: 0, valign: 'middle' });
  s.addText(f[2], { x: x + 0.3, y: 3.5, w: fw - 0.6, h: 1.05, fontFace: F, fontSize: 13,
    color: MUTED, margin: 0, lineSpacingMultiple: 1.15 });
});
card(s, M, 5.05, CW, 1.05, DARK);
s.addText('별도 서버 · API 계약 · 관리자 승인 없이, 구글 계정 권한만으로 동작합니다', {
  x: M + 0.45, y: 5.05, w: CW - 0.9, h: 1.05, fontFace: F, fontSize: 16, color: WHITE,
  margin: 0, valign: 'middle' });
s.addNotes('중요한 점은 도입 비용이 0원이라는 것입니다. 구글 스프레드시트에 내장된 기능만 사용했고, IT 부서의 서버 구축이나 외부 서비스 계약이 필요 없습니다.');

// =========================================================
// 5. 상세 기능
// =========================================================
s = pres.addSlide();
titleLight(s, '주요 기능', '규칙은 자동으로, 최종 판단은 사람이');
const feats = [
  ['자동 대상 선별', '생년월일에서 생일월을 계산하고, 재직 상태 · 직급 · 부서 규칙에 따라 대상과 비대상을 자동 분류'],
  ['개인화 메일 발송', '수신자 이름과 사용 기한을 자동으로 채워 발송하며, 발신자의 회사 서명이 그대로 첨부'],
  ['중복 발송 차단', '사원번호와 연월을 키로 발송 이력을 남겨, 같은 달 재발송을 시스템이 차단'],
  ['사람의 최종 확인', '자동 분류 결과를 체크박스로 보여주고, 발송 전 담당자가 직접 확인 · 수정']
];
const cw2 = (CW - 0.45) / 2;
feats.forEach(function (f, i) {
  const x = M + (i % 2) * (cw2 + 0.45);
  const y = 2.0 + Math.floor(i / 2) * 1.72;
  card(s, x, y, cw2, 1.5);
  badge(s, x + 0.32, y + 0.3, String(i + 1), 0.46);
  s.addText(f[0], { x: x + 0.95, y: y + 0.26, w: cw2 - 1.3, h: 0.38, fontFace: F, fontSize: 17,
    bold: true, color: INK, margin: 0, valign: 'middle' });
  s.addText(f[1], { x: x + 0.95, y: y + 0.68, w: cw2 - 1.3, h: 0.65, fontFace: F, fontSize: 13,
    color: MUTED, margin: 0 });
});
s.addText('안전장치  ·  테스트 발송으로 사전 확인   ·   발송 전 인원수 재확인   ·   실패 건은 사유와 함께 기록', {
  x: M, y: 5.6, w: CW, h: 0.4, fontFace: F, fontSize: 13, bold: true, color: GREEN, margin: 0 });
s.addNotes('자동화했지만 발송 버튼은 사람이 누릅니다. 자동 분류 결과를 체크박스로 확인한 뒤 발송하는 구조라, 예외 상황도 담당자가 조정할 수 있습니다.');

// =========================================================
// 6. 데이터 현황 (chart)
// =========================================================
s = pres.addSlide();
titleLight(s, '운영 규모', '전 직원 180명 기준 월별 안내 대상자');
s.addChart(pres.ChartType.bar, [{
  name: '안내 대상자',
  labels: ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월'],
  values: [14, 18, 5, 15, 7, 13, 11, 15, 16, 22, 14, 11]
}], {
  x: M, y: 1.95, w: 8.5, h: 3.95,
  barDir: 'col',
  chartColors: [GREEN],
  showTitle: false,
  showLegend: false,
  showValue: true,
  dataLabelPosition: 'outEnd',
  dataLabelColor: INK,
  dataLabelFontFace: F,
  dataLabelFontSize: 11,
  catAxisLabelColor: MUTED,
  catAxisLabelFontFace: F,
  catAxisLabelFontSize: 11,
  valAxisLabelColor: MUTED,
  valAxisLabelFontFace: F,
  valAxisLabelFontSize: 10,
  valAxisMaxVal: 26,
  valGridLine: { color: 'E4EAE7', size: 1 },
  catGridLine: { style: 'none' },
  barGapWidthPct: 45
});
const stats = [['180', '전 직원'], ['161', '연간 안내 대상'], ['13.4', '월평균 대상자']];
stats.forEach(function (st, i) {
  const y = 1.95 + i * 1.35;
  card(s, 9.55, y, CW - 8.8, 1.15);
  s.addText(st[0], { x: 9.8, y: y + 0.12, w: 2.4, h: 0.62, fontFace: F, fontSize: 34, bold: true,
    color: GREEN, margin: 0, valign: 'middle' });
  s.addText(st[1], { x: 9.8, y: y + 0.72, w: 2.4, h: 0.3, fontFace: F, fontSize: 12,
    color: MUTED, margin: 0 });
});
foot(s, '제외 18명(임원 · 휴직 · 제외 부서) · 이메일 미등록 1명 — 규칙에 따라 자동 제외됨');
s.addNotes('월별 편차가 큽니다. 3월은 5명, 10월은 22명입니다. 사람이 매번 세는 것보다 자동 집계가 안정적인 이유입니다.');

// =========================================================
// 7. 지속가능성
// =========================================================
s = pres.addSlide();
titleLight(s, '지속가능한 운영 구조', '후임자가 코드를 몰라도 계속 쓸 수 있도록 설계했습니다');
const sus = [
  ['설정 시트', '메일 제목 · 본문 · 제외 규칙 · 발신자명을 스프레드시트에서 직접 수정.\n코드를 열 필요가 없습니다.'],
  ['사용법 시트', '매월 작업 순서와 문제 해결 방법을 스프레드시트 안에 문서로 내장.\n별도 매뉴얼 관리가 불필요합니다.'],
  ['인수인계 절차', '이 도구는 실행하는 사람의 계정으로 발송됩니다.\n담당자 변경 시 소유권 이전과 재승인 절차를 문서에 명시했습니다.']
];
sus.forEach(function (u, i) {
  const y = 1.95 + i * 1.42;
  card(s, M, y, CW, 1.22);
  badge(s, M + 0.35, y + 0.36, String(i + 1), 0.5);
  s.addText(u[0], { x: M + 1.05, y: y + 0.22, w: 2.7, h: 0.42, fontFace: F, fontSize: 19,
    bold: true, color: INK, margin: 0, valign: 'middle' });
  s.addText(u[1], { x: M + 3.95, y: y + 0.18, w: CW - 4.3, h: 0.88, fontFace: F, fontSize: 13.5,
    color: MUTED, margin: 0, valign: 'middle' });
});
card(s, M, 6.25, CW, 0.72, DARK);
s.addText('매월 ERP 최신본만 붙여넣으면, 휴직 · 퇴직 · 신규 입사자가 자동으로 반영됩니다', {
  x: M + 0.45, y: 6.25, w: CW - 0.9, h: 0.72, fontFace: F, fontSize: 15, color: WHITE,
  margin: 0, valign: 'middle' });
s.addNotes('이 부분이 이번 작업에서 가장 신경 쓴 지점입니다. 인턴이 만든 도구는 인턴이 떠나면 멈추기 쉬운데, 그렇게 되지 않도록 설정과 문서를 도구 안에 넣었습니다.');

// =========================================================
// 8. 기대효과
// =========================================================
s = pres.addSlide();
titleLight(s, '기대 효과', '');
const eff = [
  ['월 1.5~2시간', '5분 이내', '작업 시간'],
  ['수작업 대조', '규칙 기반 자동 판별', '대상자 선별'],
  ['담당자 기억에 의존', '문서 · 설정 내장', '업무 인수인계']
];
const ew = (CW - 0.45 * 2) / 3;
eff.forEach(function (e, i) {
  const x = M + i * (ew + 0.45);
  card(s, x, 1.9, ew, 2.6);
  s.addText(e[2], { x: x + 0.32, y: 2.15, w: ew - 0.64, h: 0.32, fontFace: F, fontSize: 12,
    bold: true, color: MUTED, charSpacing: 1, margin: 0 });
  s.addText(e[0], { x: x + 0.32, y: 2.58, w: ew - 0.64, h: 0.5, fontFace: F, fontSize: 15,
    color: '96A6A0', margin: 0, valign: 'middle', strike: true });
  s.addText('↓', { x: x + 0.32, y: 3.08, w: ew - 0.64, h: 0.32, fontFace: F, fontSize: 14,
    color: GREEN, margin: 0 });
  s.addText(e[1], { x: x + 0.32, y: 3.42, w: ew - 0.64, h: 0.85, fontFace: F, fontSize: 21,
    bold: true, color: INK, margin: 0, valign: 'middle' });
});
card(s, M, 4.85, CW, 1.5, TINT);
s.addText('확장 가능성', { x: M + 0.45, y: 5.0, w: 3.0, h: 0.35, fontFace: F, fontSize: 14,
  bold: true, color: GREEN, margin: 0 });
s.addText('동일한 구조를 경조사 안내, 근속 축하, 교육 이수 리마인드 등 피플실의 다른 정기 안내 업무에 그대로 적용할 수 있습니다.',
  { x: M + 0.45, y: 5.38, w: CW - 0.9, h: 0.8, fontFace: F, fontSize: 15, color: INK, margin: 0 });
foot(s, '※ 작업 시간은 대상자 13~22명 기준 추정치');
s.addNotes('시간 절감도 크지만, 더 중요한 것은 담당자가 바뀌어도 같은 품질이 유지된다는 점입니다.');

// =========================================================
// 9. 향후 계획 (dark)
// =========================================================
s = pres.addSlide();
s.background = { color: DARK };
s.addShape(pres.ShapeType.ellipse, { x: 11.2, y: -1.2, w: 3.8, h: 3.8, fill: { color: '173A2E' } });
s.addText('향후 계획', { x: M, y: 0.75, w: CW, h: 0.8, fontFace: F, fontSize: 38, bold: true,
  color: WHITE, margin: 0, valign: 'middle' });
const plans = [
  ['회사 로고 서명 적용', 'Gmail 서명 연동을 완료해 브랜드 서명이 자동 첨부되도록 설정'],
  ['운영 계정 이관 검토', '개인 계정 종속을 없애기 위해 팀 공용 계정 또는 담당자 계정으로 이전'],
  ['타 업무 확대 적용', '경조사 · 근속 축하 등 반복 안내 업무에 동일 구조 적용']
];
plans.forEach(function (p, i) {
  const y = 2.0 + i * 1.25;
  s.addShape(pres.ShapeType.roundRect, { x: M, y: y, w: CW, h: 1.05, rectRadius: 0.08,
    fill: { color: '15352A' }, line: { color: '15352A', width: 0 } });
  badge(s, M + 0.35, y + 0.28, String(i + 1), 0.48);
  s.addText(p[0], { x: M + 1.05, y: y + 0.14, w: 3.6, h: 0.4, fontFace: F, fontSize: 17,
    bold: true, color: WHITE, margin: 0, valign: 'middle' });
  s.addText(p[1], { x: M + 4.85, y: y + 0.14, w: CW - 5.2, h: 0.77, fontFace: F, fontSize: 13.5,
    color: 'A9C3B6', margin: 0, valign: 'middle' });
});
s.addText('감사합니다', { x: M, y: 6.2, w: 6, h: 0.5, fontFace: F, fontSize: 22, bold: true,
  color: GREEN, margin: 0 });
s.addNotes('질문 받겠습니다.');

pres.writeFile({ fileName: '/tmp/claude-0/-home-user--/9a88c8de-cbf8-5c56-b4ec-d42f5ef3f5c1/scratchpad/생일휴가_자동화_PT.pptx' })
  .then(function (f) { console.log('created:', f); });
