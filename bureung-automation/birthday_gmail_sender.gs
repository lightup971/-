/**
 * 생일휴가 안내 – 사원명부(구글 스프레드시트) → Gmail 자동 발송
 * ---------------------------------------------------------------
 * ERP에서 내려받은 "사원명부" 엑셀을 그대로 붙여넣기만 하면 됩니다.
 * 생년월일에서 생일월을 자동 계산하므로, 이메일/생일월을 따로 입력할 필요가 없습니다.
 * 별도 API/토큰/외부 서버 없이 구글 내장 GmailApp으로 발송하며,
 * 최초 1회 본인 구글 계정 권한 승인만 하면 됩니다(관리자 승인 불필요).
 *
 * [설치]
 *  1) 사원명부 export를 붙여넣은 스프레드시트 → 확장 프로그램 → Apps Script
 *  2) 이 코드 전체를 붙여넣고 저장 → 스프레드시트 새로고침
 *  3) 상단 "🎂 생일휴가 안내" 메뉴 사용
 *
 * [인식하는 컬럼] 1행 머리글 기준(사원명부 export 그대로면 자동 인식)
 *  성명 · 생년월일 · 이메일 · 재직여부 · 사원번호 · 직급/직책/직위 · 부서
 *  - 생일월은 생년월일에서 자동 계산(생년월일이 없고 '생일월' 컬럼이 있으면 그걸 사용)
 */

// ===== 설정 =====
const CONFIG = {
  SHEET_NAME: '',                    // 비우면 활성 시트. 특정 시트명 지정 가능(예: '사원명부(재직, 휴직)')
  SENDER_NAME: '부릉 피플실',         // 받는 사람에게 보이는 발신자 이름
  SUBJECT: '[부릉] {월}월 생일 축하 & 생일휴가 안내 🎂',
  INCLUDE_STATUSES: ['재직'],         // 발송 대상 재직여부. 휴직자도 보내려면 ['재직','휴직']
  USE_GIVEN_NAME_ONLY: true,         // true면 인사말에 성을 뗀 이름 사용(김민희 → 민희님)
  RESEND: false,                     // true면 같은 달 이미 보낸 사람도 재발송

  // 생일휴가 제외 대상(사내 규정)
  EXCLUDE_RANKS: ['CEO','CTO','대표이사','부사장','전무','상무','이사','LV.8'], // 직급/직책/직위
  EXCLUDE_DEPTS: ['장애인고용']       // 부서
};

// 본문 템플릿 ({이름} {월} {말일} {서명} 자동 치환)
// ※ Apps Script 자동발송은 Gmail 기본 서명이 붙지 않으므로 본문 서명이 필요합니다.
const BODY_TEXT =
`안녕하세요, {이름}님.
부릉 피플실입니다.

{월}월 생일을 진심으로 축하드립니다! 🎉
생일 당월을 더욱 행복하게 보내실 수 있도록,
부릉에서는 생일휴가 1일을 드리고 있습니다.

📅 생일휴가 1일 부여
- 신청·사용 기한: 생일 당월인 {월}월 말일({말일})까지
- 말일 이후 자동 소멸되며, 당월 내에서만 사용 가능합니다.

📌 신청 방법
- 옴니이솔 → 인사관리 → My HR(ESS) → 근태신청(NEW) → [휴가신청서]
  → 휴가종류: 기타휴가 > 생일휴가 선택

📢 결재 라인
- 1차 조직장(결재) → 피플실(합의)
- ※ 반드시 '피플실(합의)'을 지정해 주세요.

뜻깊고 즐거운 생일 보내시길 바랍니다.
감사합니다.{서명}`;

// 이메일 서명(담당자 명의로 쓰려면 아래를 수정)
const SIGNATURE =
`

──────────────
주식회사 부릉(VROONG) People실`;

const LOG_SHEET = '_발송기록';   // 중복발송 방지용(자동 생성)


// ===== 메뉴 =====
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎂 생일휴가 안내')
    .addItem('① 이번 달 대상자 미리보기', 'previewTargets')
    .addItem('② 나에게 테스트 발송', 'sendTestToMe')
    .addSeparator()
    .addItem('③ 이번 달 생일자에게 발송', 'sendBirthdayEmails')
    .addSeparator()
    .addItem('발송기록 초기화', 'resetLog')
    .addItem('남은 일일 발송량 확인', 'showQuota')
    .addToUi();
}


// ===== 유틸 =====
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = CONFIG.SHEET_NAME ? ss.getSheetByName(CONFIG.SHEET_NAME) : ss.getActiveSheet();
  if (!sh) throw new Error('시트를 찾을 수 없습니다: ' + CONFIG.SHEET_NAME);
  return sh;
}

function lastDayOfMonth_(year, month) { return new Date(year, month, 0).getDate(); }

// 정확히 일치하는 머리글을 우선, 없으면 부분일치
function findCol_(headers, candidates) {
  const norm = headers.map(h => String(h).trim());
  for (const c of candidates) {
    const i = norm.findIndex(h => h === c);
    if (i >= 0) return i;
  }
  for (const c of candidates) {
    const i = norm.findIndex(h => h.toLowerCase().includes(c.toLowerCase()));
    if (i >= 0) return i;
  }
  return -1;
}

function monthFromBirth_(v) {
  if (v == null || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return v.getMonth() + 1;
  const m = String(v).match(/^\s*\d{4}[-./](\d{1,2})/);
  if (m) return parseInt(m[1], 10);
  const m2 = String(v).match(/^\s*(\d{1,2})\s*$/); // '생일월' 컬럼(숫자)일 때
  return m2 ? parseInt(m2[1], 10) : null;
}

function isValidEmail_(e) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(e || '').trim()); }

function greetingName_(fullName) {
  const n = String(fullName || '').trim();
  if (!CONFIG.USE_GIVEN_NAME_ONLY) return n;
  return n.length > 1 ? n.slice(1) : n;   // 성 1자 가정(대부분의 경우)
}


// 시트 파싱 → 컬럼 인덱스
function parse_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) throw new Error('데이터가 없습니다. 1행 머리글 아래에 사원명부를 붙여넣으세요.');
  const headers = values[0];
  const col = {
    empno:  findCol_(headers, ['사원번호','사번','employeeid']),
    name:   findCol_(headers, ['성명','이름','name']),
    birth:  findCol_(headers, ['생년월일','생일','출생','birth','dob']),
    bmonth: findCol_(headers, ['생일월']),
    email:  findCol_(headers, ['이메일','메일','email','gmail']),
    status: findCol_(headers, ['재직여부','재직상태','status']),
    dept:   findCol_(headers, ['부서']),   // '상위부서'보다 '부서' 정확일치 우선
    rank:   findCol_(headers, ['직급']),
    role:   findCol_(headers, ['직책']),
    title:  findCol_(headers, ['직위'])
  };
  if (col.email < 0) throw new Error("'이메일' 컬럼을 찾을 수 없습니다.");
  if (col.birth < 0 && col.bmonth < 0) throw new Error("'생년월일' 또는 '생일월' 컬럼이 필요합니다.");
  return { sheet, values, col };
}

function excludeReason_(row, col) {
  if (col.status >= 0) {
    const st = String(row[col.status] || '').trim();
    if (st && CONFIG.INCLUDE_STATUSES.indexOf(st) < 0) return st || '재직 아님';
  }
  for (const ci of [col.rank, col.role, col.title]) {
    if (ci >= 0) {
      const v = String(row[ci] || '').trim();
      if (v && CONFIG.EXCLUDE_RANKS.indexOf(v) >= 0) return '임원 제외';
    }
  }
  if (col.dept >= 0) {
    const d = String(row[col.dept] || '').trim();
    if (CONFIG.EXCLUDE_DEPTS.indexOf(d) >= 0) return '부서 제외(' + d + ')';
  }
  return null;
}

// 대상 월의 발송 대상/제외/무효 분류
function classify_(targetMonth) {
  const { sheet, values, col } = parse_();
  const send = [], excluded = [], noEmail = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (col.name >= 0 && !String(row[col.name] || '').trim()) continue;
    const bm = col.birth >= 0 ? monthFromBirth_(row[col.birth]) : monthFromBirth_(row[col.bmonth]);
    if (bm !== targetMonth) continue;

    const name = col.name >= 0 ? String(row[col.name]).trim() : '';
    const empno = col.empno >= 0 ? String(row[col.empno] || '').trim() : ('row' + r);
    const email = String(row[col.email] || '').trim();
    const reason = excludeReason_(row, col);
    const rec = { empno, name, email };
    if (reason) { rec.reason = reason; excluded.push(rec); }
    else if (!isValidEmail_(email)) noEmail.push(rec);
    else send.push(rec);
  }
  return { send, excluded, noEmail };
}

function buildMessage_(fullName, month) {
  const year = new Date().getFullYear();
  const lastLabel = month + '월 ' + lastDayOfMonth_(year, month) + '일';
  const rep = s => String(s)
    .replace(/{이름}/g, greetingName_(fullName))
    .replace(/{월}/g, month)
    .replace(/{말일}/g, lastLabel)
    .replace(/{서명}/g, SIGNATURE);
  const text = rep(BODY_TEXT);
  return { subject: rep(CONFIG.SUBJECT), text, html: textToHtml_(text) };
}

function textToHtml_(text) {
  const esc = s => s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const body = text.split('\n').map(line => {
    const t = line.trim();
    if (/^[📅📌📢]/.test(t)) return '<div style="margin:16px 0 6px;font-weight:600;">'+esc(line)+'</div>';
    if (t === '') return '<div style="height:8px;"></div>';
    if (t.startsWith('- ')) return '<div style="padding-left:14px;">• '+esc(t.slice(2))+'</div>';
    return '<div>'+esc(line)+'</div>';
  }).join('');
  return '<div style="font-family:\'Apple SD Gothic Neo\',\'Malgun Gothic\',sans-serif;font-size:15px;line-height:1.7;color:#1a1a18;max-width:560px;">'+body+'</div>';
}


// ===== 중복발송 로그(사원번호|YYYYMM) =====
function logKey_(empno, year, month) { return empno + '|' + year + ('0'+month).slice(-2); }

function getLog_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) { sh = ss.insertSheet(LOG_SHEET); sh.appendRow(['키','사원번호','성명','이메일','발송일시']); sh.hideSheet(); }
  return sh;
}
function loadSentKeys_() {
  const sh = getLog_();
  const v = sh.getDataRange().getValues();
  const set = {};
  for (let i = 1; i < v.length; i++) set[String(v[i][0])] = true;
  return set;
}


// ===== 미리보기 / 테스트 =====
function previewTargets() {
  const ui = SpreadsheetApp.getUi();
  const month = new Date().getMonth() + 1;
  const { send, excluded, noEmail } = classify_(month);
  const names = a => a.map(x => x.name + (x.reason ? '('+x.reason+')' : '')).join(', ') || '없음';
  const sample = send.length ? buildMessage_(send[0].name, month) : null;
  let msg = month + '월 분류\n';
  msg += '- 발송 대상 ' + send.length + '명: ' + names(send) + '\n';
  msg += '- 제외 ' + excluded.length + '명: ' + names(excluded) + '\n';
  msg += '- 이메일 없음 ' + noEmail.length + '명: ' + names(noEmail) + '\n';
  if (sample) msg += '\n[예시 메일 · ' + send[0].name + ']\n제목: ' + sample.subject + '\n\n' + sample.text;
  ui.alert(msg);
}

function sendTestToMe() {
  const ui = SpreadsheetApp.getUi();
  const me = Session.getActiveUser().getEmail();
  if (!me) { ui.alert('본인 이메일을 확인할 수 없습니다.'); return; }
  const month = new Date().getMonth() + 1;
  const m = buildMessage_('테스트', month);
  GmailApp.sendEmail(me, '[테스트] ' + m.subject, m.text, { htmlBody: m.html, name: CONFIG.SENDER_NAME });
  ui.alert('테스트 메일을 ' + me + ' 로 보냈습니다.');
}


// ===== 실제 발송 =====
function sendBirthdayEmails() {
  const ui = SpreadsheetApp.getUi();
  const nowMonth = new Date().getMonth() + 1;
  const resp = ui.prompt('생일휴가 안내 발송',
    '발송할 대상 월(1~12)을 입력하세요.\n기본값: ' + nowMonth + '월', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const input = resp.getResponseText().trim();
  const month = input ? parseInt(input, 10) : nowMonth;
  if (!(month >= 1 && month <= 12)) { ui.alert('1~12 사이의 월을 입력하세요.'); return; }

  const year = new Date().getFullYear();
  const { send, excluded, noEmail } = classify_(month);
  const sentKeys = CONFIG.RESEND ? {} : loadSentKeys_();
  const toSend = send.filter(t => !sentKeys[logKey_(t.empno, year, month)]);
  const dup = send.length - toSend.length;

  if (!toSend.length) {
    ui.alert(month + '월 발송할 대상이 없습니다.\n(대상 ' + send.length + ' · 이미발송 ' + dup +
             ' · 제외 ' + excluded.length + ' · 이메일없음 ' + noEmail.length + ')');
    return;
  }
  const quota = MailApp.getRemainingDailyQuota();
  if (quota < toSend.length) { ui.alert('일일 발송 한도 부족: 대상 ' + toSend.length + ' / 남은 ' + quota + '통'); return; }

  const confirm = ui.alert('발송 확인',
    month + '월 생일자 ' + toSend.length + '명에게 발송합니다.\n' +
    '(제외 ' + excluded.length + ' · 이미발송 ' + dup + ' · 이메일없음 ' + noEmail.length + ')\n\n계속할까요?',
    ui.ButtonSet.OK_CANCEL);
  if (confirm !== ui.Button.OK) return;

  const log = getLog_();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  let ok = 0, fail = 0; const fails = [];
  toSend.forEach(t => {
    try {
      const m = buildMessage_(t.name, month);
      GmailApp.sendEmail(t.email, m.subject, m.text, { htmlBody: m.html, name: CONFIG.SENDER_NAME });
      log.appendRow([logKey_(t.empno, year, month), t.empno, t.name, t.email, stamp]);
      ok++;
    } catch (e) { fail++; fails.push(t.name + ': ' + e.message); }
    Utilities.sleep(300);
  });

  let done = '발송 완료\n성공 ' + ok + '명 · 실패 ' + fail + '명\n남은 일일 한도 약 ' + MailApp.getRemainingDailyQuota() + '통';
  if (fails.length) done += '\n\n[실패]\n' + fails.join('\n');
  if (noEmail.length) done += '\n\n[이메일 없음 · 수동 확인]\n' + noEmail.map(x => x.name).join(', ');
  ui.alert(done);
}


// ===== 관리 =====
function resetLog() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('발송기록 초기화', LOG_SHEET + ' 기록을 모두 지웁니다(중복방지 초기화). 계속할까요?', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const sh = getLog_();
  sh.clearContents(); sh.appendRow(['키','사원번호','성명','이메일','발송일시']);
  ui.alert('초기화했습니다.');
}

function showQuota() {
  SpreadsheetApp.getUi().alert('오늘 남은 Gmail 발송 한도: 약 ' + MailApp.getRemainingDailyQuota() +
    '통\n(개인 500 / 워크스페이스 1,500 기준, 다음 날 초기화)');
}
