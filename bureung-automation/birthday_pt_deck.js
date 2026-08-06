const pptxgen = require('pptxgenjs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';           // 13.33 x 7.5
pres.author = '피플실';
pres.title = '생일휴가 안내 업무 자동화';

// ===== Palette (VROONG brand green forward) =====
const DARK  = '0E2A22';
const GREEN = '00C878';
const WHITE = 'FFFFFF';
const TINT  = 'F1F6F3';
const MUTED = '5F7069';
const INK   = '15211D';
const WARN  = 'C7502F';

const F = 'Calibri';
const W = 13.33, H = 7.5;
const M = 0.75;
const CW = W - M * 2;

function titleLight(s, txt, sub) {
  s.addText(txt, { x: M, y: 0.45, w: CW, h: 0.72, fontFace: F, fontSize: 36, bold: true,
                   color: INK, margin: 0, valign: 'middle' });
  if (sub) s.addText(sub, { x: M, y: 1.16, w: CW, h: 0.38, fontFace: F, fontSize: 14,
                            color: MUTED, margin: 0, valign: 'middle' });
}
function badge(s, x, y, label, d) {
  const dia = d || 0.48;
  s.addShape(pres.ShapeType.ellipse, { x: x, y: y, w: dia, h: dia, fill: { color: GREEN } });
  s.addText(label, { x: x, y: y, w: dia, h: dia, fontFace: F, fontSize: dia > 0.55 ? 16 : 13,
                     bold: true, color: WHITE, align: 'center', valign: 'middle', margin: 0 });
}
function card(s, x, y, w, h, fill) {
  s.addShape(pres.ShapeType.roundRect, { x: x, y: y, w: w, h: h, rectRadius: 0.08,
    fill: { color: fill || TINT }, line: { color: fill || TINT, width: 0 } });
}
function foot(s, txt) {
  s.addText(txt, { x: M, y: H - 0.46, w: CW, h: 0.28, fontFace: F, fontSize: 10,
                   color: MUTED, margin: 0, align: 'left' });
}

// =========================================================
// 1. 표지
// =========================================================
let s = pres.addSlide();
s.background = { color: DARK };
s.addShape(pres.ShapeType.ellipse, { x: 10.7, y: -1.5, w: 4.6, h: 4.6, fill: { color: '173A2E' } });
s.addShape(pres.ShapeType.ellipse, { x: 11.9, y: 5.1, w: 2.6, h: 2.6, fill: { color: '173A2E' } });
s.addText('PEOPLE OPS AUTOMATION', { x: M, y: 2.05, w: 8.5, h: 0.35, fontFace: F, fontSize: 12,
  bold: true, color: GREEN, charSpacing: 3, margin: 0 });
s.addText('생일휴가 안내 업무 자동화', { x: M, y: 2.5, w: 9.6, h: 1.0, fontFace: F, fontSize: 46,
  bold: true, color: WHITE, margin: 0, valign: 'middle' });
s.addText('구글 시트 · Apps Script · 지메일 기반  |  별도 비용 없음', {
  x: M, y: 3.55, w: 9.6, h: 0.45, fontFace: F, fontSize: 17, color: 'B8CFC4', margin: 0 });
s.addShape(pres.ShapeType.roundRect, { x: M, y: 4.25, w: 3.15, h: 0.5, rectRadius: 0.1,
  fill: { color: GREEN } });
s.addText('2026년 8월 운영 완료', { x: M, y: 4.25, w: 3.15, h: 0.5, fontFace: F, fontSize: 14,
  bold: true, color: '06301F', align: 'center', valign: 'middle', margin: 0 });
s.addText('피플실 인턴 김민희  |  2026. 08. 07', { x: M, y: 5.05, w: 7, h: 0.35, fontFace: F,
  fontSize: 14, color: '8FA89B', margin: 0 });
s.addNotes('생일휴가 안내는 매월 반복되는 정형 업무입니다. 이번 발표는 이 업무를 자동화한 과정과, 담당자가 바뀌어도 계속 쓸 수 있게 만든 구조를 공유하는 자리입니다.');

// =========================================================
// 2. 추진 목적
// =========================================================
s = pres.addSlide();
titleLight(s, '추진 목적', '반복 업무를 줄이고, 사람이 바뀌어도 유지되는 구조를 만든다');
const goals = [
  ['1', '반복 업무 자동화', '매월 반복되는 명단 확인 · 대상자 선별 · 개별 발송을 자동화'],
  ['2', '누락과 실수 제거', '규칙 기반 대상자 선별, 오류 · 실수 제거'],
  ['3', '인수인계 및 지속가능성 확보', '담당자 변경 시에도 동일 품질이 유지되도록 기본설정을 도구에 내장 (GAS 코드 수정 불필요)']
];
goals.forEach(function (g, i) {
  const y = 1.9 + i * 1.42;
  card(s, M, y, CW, 1.2);
  badge(s, M + 0.34, y + 0.36, g[0]);
  s.addText(g[1], { x: M + 1.02, y: y + 0.2, w: 4.1, h: 0.8, fontFace: F, fontSize: 19,
    bold: true, color: INK, margin: 0, valign: 'middle' });
  s.addText(g[2], { x: M + 5.25, y: y + 0.18, w: CW - 5.6, h: 0.84, fontFace: F, fontSize: 14,
    color: MUTED, margin: 0, valign: 'middle' });
});
foot(s, '대상 규모: 전 직원 182명 · 연간 안내 대상 162명 · 월평균 13.5명 (2026-08-06 기준)');
s.addNotes('목적은 세 가지입니다. 시간을 줄이고, 실수를 없애고, 사람이 바뀌어도 굴러가게 하는 것. 특히 세 번째가 이번 작업의 핵심입니다.');

// =========================================================
// 3. 기존 프로세스
// =========================================================
s = pres.addSlide();
titleLight(s, '기존 프로세스', '8월 안내 기준 — 7월 말에 대상자를 확정하고 발송하는 4단계 수작업');
const procs = [
  ['0', '연 1회 · 26년 4월 말',
   'ERP 이관 완료 후 26-04-30 기준 재직자 생일 확인 → 옴니이솔 > 근태일수등록 > 사원별일수등록에서 생일휴가 1일 일괄 부여 (시작일 = 생일월 1일 / 종료일 = 생일월 말일)'],
  ['1', '매월 말 · 대상자 재확인',
   '사원명부 등 인사기록 재조회 → 신규 입사자 중 익월 생일자 추가 부여, 중도퇴사 · 휴직자 부여 내역 삭제 후 대상 제외 / 임원 부여 내역 삭제 후 대상 제외 (26-04-30 당시 제외 누락)'],
  ['2', '검토', '추려진 대상자 목록을 사수님께 전달해 컨펌'],
  ['3', '발송', '슬랙에서 대상자 조회 (동명이인 주의 필요) → 템플릿의 "월"과 "이름"을 직접 수정해 1명씩 수동 발송']
];
procs.forEach(function (p, i) {
  const y = 1.82 + i * 0.98;
  card(s, M, y, CW, 0.88);
  badge(s, M + 0.3, y + 0.2, p[0], 0.44);
  s.addText('STEP ' + p[0], { x: M + 0.92, y: y + 0.08, w: 1.35, h: 0.3, fontFace: F,
    fontSize: 13, bold: true, color: INK, margin: 0 });
  s.addText(p[1], { x: M + 0.92, y: y + 0.38, w: 1.9, h: 0.42, fontFace: F, fontSize: 11,
    color: GREEN, bold: true, margin: 0 });
  s.addText(p[2], { x: M + 3.0, y: y + 0.06, w: CW - 3.35, h: 0.76, fontFace: F, fontSize: 12.5,
    color: MUTED, margin: 0, valign: 'middle' });
});
card(s, M, 5.82, CW, 1.28, 'FBEDE9');
s.addShape(pres.ShapeType.ellipse, { x: M + 0.3, y: 6.02, w: 0.36, h: 0.36, fill: { color: WARN } });
s.addText('!', { x: M + 0.3, y: 6.02, w: 0.36, h: 0.36, fontFace: F, fontSize: 13, bold: true,
  color: WHITE, align: 'center', valign: 'middle', margin: 0 });
s.addText('문제점', { x: M + 0.86, y: 6.0, w: 1.2, h: 0.34, fontFace: F, fontSize: 14,
  bold: true, color: '8A3A22', margin: 0, valign: 'middle' });
const probs = [
  '단순 반복 누적 → 템플릿 이름 미수정 발송 사례 발생',
  '매월 동일 작업 반복에 따른 시간 소요',
  '절차 미문서화로 인수인계 시 비효율 발생',
  '슬랙 발송 시 감사 인사 등 회신 누적'
];
probs.forEach(function (t, i) {
  const x = M + 0.86 + (i % 2) * ((CW - 1.2) / 2);
  const y = 6.36 + Math.floor(i / 2) * 0.34;
  s.addText('· ' + t, { x: x, y: y, w: (CW - 1.2) / 2 - 0.1, h: 0.32, fontFace: F,
    fontSize: 11.5, color: '8A3A22', margin: 0, valign: 'middle' });
});
s.addNotes('연 1회 일괄 부여를 하고 매월 말에 변동분만 조정하는 구조입니다. 문제는 STEP 1의 대조와 STEP 3의 수동 발송이었습니다. 실제로 이름을 바꾸지 않고 보낸 적이 있었고, 다행히 상대가 확인하기 전에 수정했습니다.');

// =========================================================
// 4. 접근 방식의 전환
// =========================================================
s = pres.addSlide();
titleLight(s, '접근 방식의 전환', '처음 시도한 방식이 막히면서, 판단 기준을 바꿨습니다');
const half = (CW - 0.45) / 2;
card(s, M, 1.9, half, 3.2, 'FBEDE9');
s.addShape(pres.ShapeType.ellipse, { x: M + 0.35, y: 2.2, w: 0.46, h: 0.46, fill: { color: WARN } });
s.addText('X', { x: M + 0.35, y: 2.2, w: 0.46, h: 0.46, fontFace: F, fontSize: 15, bold: true,
  color: WHITE, align: 'center', valign: 'middle', margin: 0 });
s.addText('1차 시도 — 슬랙 자동화', { x: M + 0.35, y: 2.8, w: half - 0.7, h: 0.4, fontFace: F,
  fontSize: 19, bold: true, color: INK, margin: 0 });
[
  'ERP 사원명부만 넣으면 대상자 선별 후 슬랙 DM 발송하는 도구 구상',
  '인턴에게 슬랙 API 권한 미부여로 무산',
  '권한 확보하더라도 인턴 교체 시마다 재부여 필요 → 지속가능성 부족'
].forEach(function (t, i) {
  s.addText('· ' + t, { x: M + 0.35, y: 3.28 + i * 0.58, w: half - 0.7, h: 0.55, fontFace: F,
    fontSize: 13, color: MUTED, margin: 0 });
});
const rx = M + half + 0.45;
card(s, rx, 1.9, half, 3.2, TINT);
s.addShape(pres.ShapeType.ellipse, { x: rx + 0.35, y: 2.2, w: 0.46, h: 0.46, fill: { color: GREEN } });
s.addText('v', { x: rx + 0.35, y: 2.2, w: 0.46, h: 0.46, fontFace: F, fontSize: 15, bold: true,
  color: WHITE, align: 'center', valign: 'middle', margin: 0 });
s.addText('전환 — 구글 시트 + Gmail', { x: rx + 0.35, y: 2.8, w: half - 0.7, h: 0.4, fontFace: F,
  fontSize: 19, bold: true, color: INK, margin: 0 });
[
  '관리자 승인 없이 본인 구글 계정 권한만으로 동작',
  '담당자 변경 시 동일 절차로 인수 가능',
  '추가 비용 · 외부 서비스 계약 없음'
].forEach(function (t, i) {
  s.addText('· ' + t, { x: rx + 0.35, y: 3.28 + i * 0.58, w: half - 0.7, h: 0.55, fontFace: F,
    fontSize: 13, color: MUTED, margin: 0 });
});
card(s, M, 5.42, CW, 1.05, DARK);
s.addText('판단 기준:  "권한 없이 동작하는가"  +  "다음 담당자가 이어받을 수 있는가"', {
  x: M + 0.45, y: 5.42, w: CW - 0.9, h: 1.05, fontFace: F, fontSize: 16, color: WHITE,
  margin: 0, valign: 'middle' });
s.addNotes('슬랙 API 권한이 막힌 것이 오히려 전환점이 됐습니다. 권한을 받아냈더라도 인턴이 바뀔 때마다 다시 받아야 했을 텐데, 그건 지속가능한 구조가 아니라고 판단했습니다.');

// =========================================================
// 5. 개선 방식
// =========================================================
s = pres.addSlide();
titleLight(s, '개선 방식', '스프레드시트 메뉴 4단계로 완료됩니다');
const flow = [
  ['0', '명단 대조', 'ERP 최신본과\n발송명단 자동 비교\n\n재직상태 · 직급 변동 반영\n신규 입사자 자동 추가\n테스트 계정 자동 판별'],
  ['1', '대상 분류', '생일월 자동 계산\n\n임원 · 휴직 · 제외 부서\n(장애인고용) 판별 후\n대상자에서 제외'],
  ['2', '테스트 발송', '담당자 본인 메일로\n견본 발송\n\n메일 제목 · 안내 문구\n서명 적용 여부 사전 확인'],
  ['3', '일괄 발송', '개인화된 안내 메일 발송\n\n결과 자동 기록']
];
const fw = (CW - 0.35 * 3) / 4;
flow.forEach(function (f, i) {
  const x = M + i * (fw + 0.35);
  card(s, x, 1.9, fw, 3.55);
  badge(s, x + 0.3, 2.18, f[0], 0.58);
  s.addText(f[1], { x: x + 0.3, y: 2.86, w: fw - 0.6, h: 0.4, fontFace: F, fontSize: 17,
    bold: true, color: INK, margin: 0, valign: 'middle' });
  s.addText(f[2], { x: x + 0.3, y: 3.3, w: fw - 0.6, h: 2.02, fontFace: F, fontSize: 11.5,
    color: MUTED, margin: 0, lineSpacingMultiple: 1.1 });
});
card(s, M, 5.72, CW, 0.95, DARK);
s.addText('별도 서버 · API · 관리자 승인 없이 구글 계정 권한만으로 동작합니다', {
  x: M + 0.45, y: 5.72, w: CW - 0.9, h: 0.95, fontFace: F, fontSize: 16, color: WHITE,
  margin: 0, valign: 'middle' });
s.addNotes('도입 비용이 0원입니다. 구글 스프레드시트에 내장된 기능만 사용했고, IT 부서의 서버 구축이나 외부 서비스 계약이 필요 없습니다.');

// =========================================================
// 6. 주요 기능
// =========================================================
s = pres.addSlide();
titleLight(s, '주요 기능', '규칙은 자동으로, 최종 판단은 사람이');
const feats = [
  ['발송명단 자동 최신화', '재직여부 · 직급 · 직책 · 비용센터 · 이메일 5개 항목 변경 감지. 변경 내역은 [대조결과] 시트에 기록되고, ERP 양식의 열 순서가 달라도 항목명으로 자동 인식'],
  ['자동 대상 선별', '생년월일에서 생일월을 계산하고, 재직 상태 · 직급 · 부서 규칙에 따라 대상 / 비대상으로 분류'],
  ['개인화된 메일 발송', '수신자 이름과 사용 기한을 자동 치환해 발송'],
  ['중복 발송 차단', '[발송기록] 시트에서 사원번호 + 연월 기준으로 이력을 관리하고, 중복 발송 시 안내창 표시'],
  ['담당자의 최종 확인', '자동 분류 결과를 체크박스로 제시하고, 발송 전 담당자가 최종 확인 · 수정해 대상자 확정']
];
const cw2 = (CW - 0.45) / 2;
feats.forEach(function (f, i) {
  const x = M + (i % 2) * (cw2 + 0.45);
  const y = 1.88 + Math.floor(i / 2) * 1.48;
  card(s, x, y, cw2, 1.3);
  badge(s, x + 0.3, y + 0.26, String(i + 1), 0.44);
  s.addText(f[0], { x: x + 0.88, y: y + 0.2, w: cw2 - 1.2, h: 0.36, fontFace: F, fontSize: 16,
    bold: true, color: INK, margin: 0, valign: 'middle' });
  s.addText(f[1], { x: x + 0.88, y: y + 0.58, w: cw2 - 1.2, h: 0.62, fontFace: F, fontSize: 12,
    color: MUTED, margin: 0 });
});
s.addText('안전장치   ·  테스트 발송(사전 확인)    ·  발송 전 대상자 인원수 재확인    ·  실패 건 사유 기록', {
  x: M, y: 6.42, w: CW, h: 0.38, fontFace: F, fontSize: 13, bold: true, color: GREEN, margin: 0 });
s.addNotes('자동화했지만 발송 버튼은 사람이 누릅니다. 자동 분류 결과를 체크박스로 확인한 뒤 발송하는 구조라, 예외 상황도 담당자가 조정할 수 있습니다.');

// =========================================================
// 7. 운영 규모
// =========================================================
s = pres.addSlide();
titleLight(s, '운영 규모', '전 직원 182명 기준 월별 안내 대상자 (2026-08-06 기준)');
s.addChart(pres.ChartType.bar, [{
  name: '안내 대상자',
  labels: ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월'],
  values: [14, 18, 5, 15, 7, 13, 11, 15, 16, 22, 14, 11]
}], {
  x: M, y: 1.9, w: 8.5, h: 4.0,
  barDir: 'col', chartColors: [GREEN],
  showTitle: false, showLegend: false, showValue: true,
  dataLabelPosition: 'outEnd', dataLabelColor: INK,
  dataLabelFontFace: F, dataLabelFontSize: 11,
  catAxisLabelColor: MUTED, catAxisLabelFontFace: F, catAxisLabelFontSize: 11,
  valAxisLabelColor: MUTED, valAxisLabelFontFace: F, valAxisLabelFontSize: 10,
  valAxisMaxVal: 26,
  valGridLine: { color: 'E4EAE7', size: 1 },
  catGridLine: { style: 'none' },
  barGapWidthPct: 45
});
const stats = [['182', '전 직원'], ['162', '연간 안내 대상'], ['13.5', '월평균 대상자']];
stats.forEach(function (st, i) {
  const y = 1.9 + i * 1.38;
  card(s, 9.55, y, CW - 8.8, 1.18);
  s.addText(st[0], { x: 9.8, y: y + 0.12, w: 2.4, h: 0.64, fontFace: F, fontSize: 34, bold: true,
    color: GREEN, margin: 0, valign: 'middle' });
  s.addText(st[1], { x: 9.8, y: y + 0.74, w: 2.4, h: 0.3, fontFace: F, fontSize: 12,
    color: MUTED, margin: 0 });
});
foot(s, '제외 18명(임원 · 휴직 · 제외 부서) · 이메일 미등록 1명 — 규칙에 따라 자동 제외    |    3월 5명 ~ 10월 22명, 월별 편차 큼');
s.addNotes('월별 편차가 큽니다. 3월은 5명, 10월은 22명입니다. 사람이 매번 세는 것보다 자동 집계가 안정적인 이유입니다.');

// =========================================================
// 8. 8월 운영 결과
// =========================================================
s = pres.addSlide();
titleLight(s, '2026년 8월 운영 결과', '실제 발송까지 완료 — 제안이 아닌 실증');
const res = [['15', '8월 안내 대상자'], ['15', '발송 완료'], ['0', '누락 · 오발송']];
const rw = (CW - 0.45 * 2) / 3;
res.forEach(function (r, i) {
  const x = M + i * (rw + 0.45);
  card(s, x, 1.88, rw, 1.7, i === 1 ? DARK : TINT);
  s.addText(r[0], { x: x + 0.3, y: 2.08, w: rw - 0.6, h: 0.9, fontFace: F, fontSize: 50,
    bold: true, color: i === 1 ? GREEN : (i === 2 ? GREEN : INK), margin: 0, valign: 'middle' });
  s.addText(r[1], { x: x + 0.3, y: 3.0, w: rw - 0.6, h: 0.34, fontFace: F, fontSize: 14,
    color: i === 1 ? 'B8CFC4' : MUTED, margin: 0 });
});
card(s, M, 3.75, CW, 1.0, 'FBEDE9');
s.addText('기존 방식', { x: M + 0.35, y: 3.88, w: 1.3, h: 0.3, fontFace: F, fontSize: 12,
  bold: true, color: WARN, margin: 0 });
s.addText('사원명부 등 인사정보 단순 조회 시 퇴직자(김대현님) · 휴직자(소유빈님) 각 1명을 거르지 못하는 실수 발생', {
  x: M + 1.75, y: 3.75, w: CW - 2.1, h: 1.0, fontFace: F, fontSize: 13.5, color: '8A3A22',
  margin: 0, valign: 'middle' });
s.addText('이번 운영에서 확인된 것', { x: M, y: 4.95, w: CW, h: 0.36, fontFace: F, fontSize: 15,
  bold: true, color: INK, margin: 0 });
[
  '퇴직자 · 휴직자 2명이 규칙에 따라 자동 비대상 처리되어, 잘못된 발송이 발생하지 않았습니다.',
  'ERP 사원명부 기반으로 정리한 발송명단 정보만으로 15명 전원에게 발송했습니다.',
  '발송 이력이 자동 기록되어, 같은 달 중복 발송이 차단되는 것을 확인했습니다.'
].forEach(function (p, i) {
  const y = 5.38 + i * 0.55;
  s.addShape(pres.ShapeType.ellipse, { x: M + 0.05, y: y + 0.04, w: 0.32, h: 0.32, fill: { color: GREEN } });
  s.addText('v', { x: M + 0.05, y: y + 0.04, w: 0.32, h: 0.32, fontFace: F, fontSize: 12, bold: true,
    color: WHITE, align: 'center', valign: 'middle', margin: 0 });
  s.addText(p, { x: M + 0.58, y: y, w: CW - 0.65, h: 0.4, fontFace: F, fontSize: 13,
    color: INK, margin: 0, valign: 'middle' });
});
foot(s, '9월 이후 안내도 동일 절차로 진행 가능');
s.addNotes('가장 중요한 슬라이드입니다. 아이디어 단계가 아니라 8월 안내를 실제로 이 도구로 발송했고, 대상자 15명 전원에게 누락 없이 나갔습니다.');

// =========================================================
// 9. 지속가능한 운영 구조
// =========================================================
s = pres.addSlide();
titleLight(s, '지속가능한 운영 구조', '후임자가 GAS 코드를 직접 수정하지 않더라도 계속 쓸 수 있도록 설계');
const sus = [
  ['[기본설정] 시트', '메일 제목 · 본문 · 제외 규칙 · 발신자명을 시트에서 직접 수정.\n코드(Apps Script)를 열 필요가 없습니다.'],
  ['[사용법] 시트', '연간 흐름 · 매월 작업 순서 · 문제 해결 방법을 스프레드시트에 수록.\n별도 매뉴얼 관리가 불필요합니다.'],
  ['인수인계 절차 명문화', '실행자 계정으로 발송되는 구조를 명시하고,\n소유권 이전 및 권한 재승인 절차를 문서화했습니다.']
];
sus.forEach(function (u, i) {
  const y = 1.95 + i * 1.4;
  card(s, M, y, CW, 1.2);
  badge(s, M + 0.34, y + 0.36, String(i + 1), 0.48);
  s.addText(u[0], { x: M + 1.02, y: y + 0.2, w: 3.0, h: 0.8, fontFace: F, fontSize: 18,
    bold: true, color: INK, margin: 0, valign: 'middle' });
  s.addText(u[1], { x: M + 4.2, y: y + 0.16, w: CW - 4.55, h: 0.88, fontFace: F, fontSize: 13,
    color: MUTED, margin: 0, valign: 'middle' });
});
card(s, M, 6.2, CW, 0.72, DARK);
s.addText('매월 ERP 최신본만 붙여넣고 [명단 대조]를 실행하면, 직급 · 재직상태 변동과 신규 입사자가 자동 반영됩니다', {
  x: M + 0.45, y: 6.2, w: CW - 0.9, h: 0.72, fontFace: F, fontSize: 14.5, color: WHITE,
  margin: 0, valign: 'middle' });
s.addNotes('이 부분이 가장 신경 쓴 지점입니다. 인턴이 만든 도구는 인턴이 떠나면 멈추기 쉬운데, 그렇게 되지 않도록 설정과 문서를 도구 안에 넣었습니다.');

// =========================================================
// 10. 기대 효과
// =========================================================
s = pres.addSlide();
titleLight(s, '기대 효과', '시간 절감보다 중요한 것은 품질의 일관성');
const eff = [
  ['작업 시간', '20분', '5분 이내'],
  ['대상자 선별', '수작업 대조', '규칙 기반 자동 판별'],
  ['업무 인수인계', '담당자 기억 · 개인 기록에 의존', '문서 · 설정에 내장']
];
const ew = (CW - 0.45 * 2) / 3;
eff.forEach(function (e, i) {
  const x = M + i * (ew + 0.45);
  card(s, x, 1.88, ew, 2.55);
  s.addText(e[0], { x: x + 0.3, y: 2.12, w: ew - 0.6, h: 0.32, fontFace: F, fontSize: 12,
    bold: true, color: MUTED, charSpacing: 1, margin: 0 });
  s.addText(e[1], { x: x + 0.3, y: 2.54, w: ew - 0.6, h: 0.5, fontFace: F, fontSize: 14,
    color: '96A6A0', margin: 0, valign: 'middle', strike: true });
  s.addText('↓', { x: x + 0.3, y: 3.04, w: ew - 0.6, h: 0.32, fontFace: F, fontSize: 14,
    color: GREEN, margin: 0 });
  s.addText(e[2], { x: x + 0.3, y: 3.38, w: ew - 0.6, h: 0.85, fontFace: F, fontSize: 19,
    bold: true, color: INK, margin: 0, valign: 'middle' });
});
card(s, M, 4.78, CW, 1.45, TINT);
s.addText('확장 가능성', { x: M + 0.45, y: 4.92, w: 3.0, h: 0.34, fontFace: F, fontSize: 14,
  bold: true, color: GREEN, margin: 0 });
s.addText('동일한 구조를 경조사 안내 · 근속 축하 · 교육 이수 리마인드 등 반복 안내 업무에 그대로 적용할 수 있습니다.',
  { x: M + 0.45, y: 5.28, w: CW - 0.9, h: 0.8, fontFace: F, fontSize: 15, color: INK, margin: 0 });
foot(s, '정확도는 8월 운영 결과(15명 전원 발송, 누락 0)로 확인');
s.addNotes('시간 절감도 크지만, 더 중요한 것은 담당자가 바뀌어도 같은 품질이 유지된다는 점입니다.');

// =========================================================
// 11. 향후 계획 및 후속 절차
// =========================================================
s = pres.addSlide();
s.background = { color: DARK };
s.addShape(pres.ShapeType.ellipse, { x: 11.4, y: -1.3, w: 3.6, h: 3.6, fill: { color: '173A2E' } });
s.addText('향후 계획 및 후속 절차', { x: M, y: 0.6, w: CW, h: 0.75, fontFace: F, fontSize: 34,
  bold: true, color: WHITE, margin: 0, valign: 'middle' });

s.addText('향후 계획', { x: M, y: 1.6, w: half, h: 0.36, fontFace: F, fontSize: 15, bold: true,
  color: GREEN, margin: 0 });
[
  ['ERP 휴가 부여 파일 자동 생성', '현재 STEP 0 · 1의 ERP 부여 · 삭제는 수동.\n업로드 양식이 정형화되어 있어 동일 시트에서 생성 가능'],
  ['타 업무 확대 적용 가능성', '근속 축하 등 반복 안내 업무']
].forEach(function (p, i) {
  const y = 2.05 + i * 1.32;
  s.addShape(pres.ShapeType.roundRect, { x: M, y: y, w: half, h: 1.16, rectRadius: 0.08,
    fill: { color: '15352A' }, line: { color: '15352A', width: 0 } });
  badge(s, M + 0.28, y + 0.26, String(i + 1), 0.42);
  s.addText(p[0], { x: M + 0.82, y: y + 0.16, w: half - 1.1, h: 0.36, fontFace: F, fontSize: 15,
    bold: true, color: WHITE, margin: 0, valign: 'middle' });
  s.addText(p[1], { x: M + 0.82, y: y + 0.52, w: half - 1.1, h: 0.58, fontFace: F, fontSize: 11.5,
    color: 'A9C3B6', margin: 0 });
});

s.addText('후속 절차', { x: rx, y: 1.6, w: half, h: 0.36, fontFace: F, fontSize: 15, bold: true,
  color: GREEN, margin: 0 });
[
  ['소유권 이관 결정', '개인 계정 발송 구조 → 인턴 종료 시 중단.\n팀 공용 계정 또는 사수님 · 후임자 계정 이관(권한 공유) 필요'],
  ['ERP 이메일 열 노출 검토', '(인턴 기준) 신규 입사자 이메일은 [사원등록] 또는 [입퇴사자현황]에서만 조회 가능'],
  ['임원 판별 기준 확정', '현재 직급 · 직책 목록 기반 판별.\n부릉 인사규정상 "임원" 정의와 일치 여부 확인 필요']
].forEach(function (p, i) {
  const y = 2.05 + i * 1.32;
  s.addShape(pres.ShapeType.roundRect, { x: rx, y: y, w: half, h: 1.16, rectRadius: 0.08,
    fill: { color: '15352A' }, line: { color: '15352A', width: 0 } });
  badge(s, rx + 0.28, y + 0.26, String(i + 1), 0.42);
  s.addText(p[0], { x: rx + 0.82, y: y + 0.16, w: half - 1.1, h: 0.36, fontFace: F, fontSize: 15,
    bold: true, color: WHITE, margin: 0, valign: 'middle' });
  s.addText(p[1], { x: rx + 0.82, y: y + 0.52, w: half - 1.1, h: 0.58, fontFace: F, fontSize: 11.5,
    color: 'A9C3B6', margin: 0 });
});

s.addText('감사합니다', { x: M, y: 6.35, w: 6, h: 0.5, fontFace: F, fontSize: 22, bold: true,
  color: GREEN, margin: 0 });
s.addNotes('질문 받겠습니다.');

pres.writeFile({ fileName: '/tmp/claude-0/-home-user--/9a88c8de-cbf8-5c56-b4ec-d42f5ef3f5c1/scratchpad/생일휴가_자동화_PT.pptx' })
  .then(function (f) { console.log('created:', f); });
