/**
 * 생일휴가 안내 – 사원명부(구글 스프레드시트) → Gmail 발송 (검토 후 발송)
 * ---------------------------------------------------------------
 * ① 발송 대상 검토 시트 만들기  → '발송검토' 시트에 대상/제외/본문이 정리됩니다.
 * ② (시트에서 체크박스로 최종 확인·수정)
 * ③ 검토 시트 기준으로 발송     → 체크된 사람에게만, 시트에 보이는 본문 그대로 발송.
 *
 * 별도 API/토큰/외부 서버 없이 구글 내장 GmailApp으로 발송하며,
 * 최초 1회 본인 구글 계정 권한 승인만 하면 됩니다(관리자 승인 불필요).
 *
 * [인식 컬럼] 1행 머리글 기준(순서·개수 무관)
 *  성명 · 생년월일(또는 생일월) · 이메일 · 재직여부 · 사원번호 · 직급/직책/직위 · 부서/비용센터
 */

// ===== 설정 =====
const CONFIG = {
  SHEET_NAME: '',                    // 비우면 명단 시트 자동 선택. 특정 시트명 지정 가능
  SENDER_NAME: '부릉 피플실',
  SUBJECT: '[부릉] {월}월 생일 축하와 생일휴가 안내 🎂',
  INCLUDE_STATUSES: ['재직'],         // 기본 대상. 휴직 포함하려면 ['재직','휴직'] (검토시트에서 개별 체크도 가능)
  USE_GIVEN_NAME_ONLY: true,         // 인사말에 성 뗀 이름(김민희 → 민희님)
  RESEND: false,                     // true면 같은 달 이미 보낸 사람도 재발송

  EXCLUDE_RANKS: ['CEO','CTO','대표이사','부사장','전무','상무','이사','LV.8'],
  EXCLUDE_DEPTS: ['장애인고용']
};

// 본문 템플릿 ({이름} {월} {말일} {서명} 자동 치환)
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

// 이메일 서명(담당자 명의로 바꾸려면 아래 수정)
const SIGNATURE =
`

──────────────
주식회사 부릉(VROONG) People실`;

const REVIEW_SHEET = '발송검토';
const LOG_SHEET = '_발송기록';
const HEADERS = ['발송','성명','이메일','생일월','상태','사유','사원번호','제목','본문(실제 발송 내용)'];


// ===== 메뉴 =====
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎂 생일휴가 안내')
    .addItem('① 발송 대상 검토 시트 만들기', 'buildReviewSheet')
    .addItem('② 나에게 테스트 발송', 'sendTestToMe')
    .addItem('③ 검토 시트 기준으로 발송', 'sendFromReview')
    .addSeparator()
    .addItem('발송기록 초기화', 'resetLog')
    .addItem('남은 일일 발송량 확인', 'showQuota')
    .addToUi();
}


// ===== 시트/컬럼 =====
function getRosterSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (CONFIG.SHEET_NAME) {
    const s = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!s) throw new Error('명단 시트를 찾을 수 없습니다: ' + CONFIG.SHEET_NAME);
    return s;
  }
  const active = ss.getActiveSheet();
  if ([REVIEW_SHEET, LOG_SHEET].indexOf(active.getName()) < 0) return active;
  // 검토/기록 시트가 활성일 때는 그 외 첫 시트를 명단으로 사용
  const other = ss.getSheets().filter(s => [REVIEW_SHEET, LOG_SHEET].indexOf(s.getName()) < 0);
  if (!other.length) throw new Error('명단 시트를 찾을 수 없습니다.');
  return other[0];
}

function findCol_(headers, candidates, exactOnly) {
  const norm = headers.map(h => String(h).trim());
  for (const c of candidates) { const i = norm.indexOf(c); if (i >= 0) return i; }
  if (!exactOnly) for (const c of candidates) {
    const i = norm.findIndex(h => h.toLowerCase().includes(c.toLowerCase())); if (i >= 0) return i;
  }
  return -1;
}

function parse_() {
  const sheet = getRosterSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) throw new Error('명단 데이터가 없습니다. 1행 머리글 아래에 사원명부를 넣으세요.');
  const h = values[0];
  const col = {
    empno:  findCol_(h, ['사원번호','사번']),
    name:   findCol_(h, ['성명','이름','name']),
    birth:  findCol_(h, ['생년월일','생일','출생','birth','dob']),
    bmonth: findCol_(h, ['생일월'], true),
    email:  findCol_(h, ['이메일','메일','email','gmail']),
    status: findCol_(h, ['재직여부','재직상태','status']),
    dept:   findCol_(h, ['부서','비용센터']),
    rank:   findCol_(h, ['직급'], true),
    role:   findCol_(h, ['직책'], true),
    title:  findCol_(h, ['직위'], true)
  };
  if (col.email < 0) throw new Error("'이메일' 컬럼을 찾을 수 없습니다.");
  if (col.birth < 0 && col.bmonth < 0) throw new Error("'생년월일' 또는 '생일월' 컬럼이 필요합니다.");
  return { sheet, values, col };
}

function lastDayOfMonth_(y, m) { return new Date(y, m, 0).getDate(); }
function monthFromBirth_(v) {
  if (v == null || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return v.getMonth() + 1;
  let m = String(v).match(/^\s*\d{4}[-./](\d{1,2})/); if (m) return parseInt(m[1], 10);
  m = String(v).match(/^\s*(\d{1,2})\s*$/); return m ? parseInt(m[1], 10) : null;
}
function isValidEmail_(e) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(e || '').trim()); }
function greetingName_(n) { n = String(n || '').trim(); return CONFIG.USE_GIVEN_NAME_ONLY && n.length > 1 ? n.slice(1) : n; }

function excludeReason_(row, col) {
  if (col.status >= 0) {
    const st = String(row[col.status] || '').trim();
    if (st && CONFIG.INCLUDE_STATUSES.indexOf(st) < 0) return st || '재직 아님';
  }
  for (const ci of [col.rank, col.role, col.title]) if (ci >= 0) {
    const v = String(row[ci] || '').trim(); if (v && CONFIG.EXCLUDE_RANKS.indexOf(v) >= 0) return '임원 제외';
  }
  if (col.dept >= 0) { const d = String(row[col.dept] || '').trim(); if (CONFIG.EXCLUDE_DEPTS.indexOf(d) >= 0) return '부서 제외(' + d + ')'; }
  return null;
}

function classify_(targetMonth) {
  const { values, col } = parse_();
  const send = [], excluded = [], noEmail = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (col.name >= 0 && !String(row[col.name] || '').trim()) continue;
    const bm = col.birth >= 0 ? monthFromBirth_(row[col.birth]) : monthFromBirth_(row[col.bmonth]);
    if (bm !== targetMonth) continue;
    const rec = {
      name:  col.name >= 0 ? String(row[col.name]).trim() : '',
      email: String(row[col.email] || '').trim(),
      empno: col.empno >= 0 ? String(row[col.empno] || '').trim() : ''
    };
    const reason = excludeReason_(row, col);
    if (reason) { rec.reason = reason; excluded.push(rec); }
    else if (!isValidEmail_(rec.email)) noEmail.push(rec);
    else send.push(rec);
  }
  return { send, excluded, noEmail };
}

function buildMessage_(fullName, month) {
  const y = new Date().getFullYear();
  const lastLabel = month + '월 ' + lastDayOfMonth_(y, month) + '일';
  const rep = s => String(s).replace(/{이름}/g, greetingName_(fullName)).replace(/{월}/g, month)
                            .replace(/{말일}/g, lastLabel).replace(/{서명}/g, SIGNATURE);
  return { subject: rep(CONFIG.SUBJECT), text: rep(BODY_TEXT) };
}

function textToHtml_(text) {
  const esc = s => s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const body = String(text).split('\n').map(line => {
    const t = line.trim();
    if (/^[📅📌📢]/.test(t)) return '<div style="margin:16px 0 6px;font-weight:600;">'+esc(line)+'</div>';
    if (t === '') return '<div style="height:8px;"></div>';
    if (t.startsWith('- ')) return '<div style="padding-left:14px;">• '+esc(t.slice(2))+'</div>';
    return '<div>'+esc(line)+'</div>';
  }).join('');
  return '<div style="font-family:\'Apple SD Gothic Neo\',\'Malgun Gothic\',sans-serif;font-size:15px;line-height:1.7;color:#1a1a18;max-width:560px;">'+body+'</div>';
}


// ===== ① 검토 시트 만들기 =====
function buildReviewSheet() {
  const ui = SpreadsheetApp.getUi();
  const nowMonth = new Date().getMonth() + 1;
  const resp = ui.prompt('발송 대상 검토', '대상 월(1~12)을 입력하세요.\n기본값: ' + nowMonth + '월', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const input = resp.getResponseText().trim();
  const month = input ? parseInt(input, 10) : nowMonth;
  if (!(month >= 1 && month <= 12)) { ui.alert('1~12 사이로 입력하세요.'); return; }

  const { send, excluded, noEmail } = classify_(month);

  const mk = (rec, status, checked) => {
    const m = buildMessage_(rec.name, month);
    return [checked, rec.name, rec.email, month, status, rec.reason || '', rec.empno, m.subject, m.text];
  };
  const data = []
    .concat(send.map(r => mk(r, '대상', true)))
    .concat(excluded.map(r => mk(r, '제외', false)))
    .concat(noEmail.map(r => mk(r, '이메일없음', false)));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(REVIEW_SHEET);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(REVIEW_SHEET, 0);

  // 안내 타이틀
  sh.getRange(1, 1).setValue('▶ ' + month + '월 생일휴가 발송 검토 — [발송] 열을 체크한 사람에게만, [본문] 내용 그대로 발송됩니다. 확인 후 메뉴 ③ 실행.');
  sh.getRange(1, 1, 1, HEADERS.length).merge().setFontWeight('bold').setWrap(true).setBackground('#fff3cd');
  // 헤더
  sh.getRange(2, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#eeeeee');

  if (data.length) {
    sh.getRange(3, 1, data.length, HEADERS.length).setValues(data);
    const cb = sh.getRange(3, 1, data.length, 1);
    cb.insertCheckboxes();
    cb.setValues(data.map(r => [r[0]]));
  }

  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 46);   // 발송
  sh.setColumnWidth(2, 80);   // 성명
  sh.setColumnWidth(3, 210);  // 이메일
  sh.setColumnWidth(4, 60);   // 생일월
  sh.setColumnWidth(5, 80);   // 상태
  sh.setColumnWidth(6, 110);  // 사유
  sh.setColumnWidth(7, 110);  // 사원번호
  sh.setColumnWidth(8, 260);  // 제목
  sh.setColumnWidth(9, 520);  // 본문
  sh.getRange(3, 9, Math.max(data.length,1), 1).setWrap(true).setVerticalAlignment('top');
  sh.activate();

  ui.alert(month + '월 검토 시트 생성 완료\n' +
    '- 대상 ' + send.length + '명(체크됨) · 제외 ' + excluded.length + '명 · 이메일없음 ' + noEmail.length + '명\n\n' +
    "'발송검토' 시트에서 체크·본문을 확인/수정한 뒤, 메뉴 ③으로 발송하세요.\n" +
    '(제외된 사람도 체크하면 발송됩니다. 본문도 직접 수정 가능합니다.)');
}


// ===== ③ 검토 시트 기준 발송 =====
function idxOf_(header, name) { return header.indexOf(name); }

function sendFromReview() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(REVIEW_SHEET);
  if (!sh) { ui.alert("먼저 메뉴 ①로 '발송검토' 시트를 만들어주세요."); return; }

  const values = sh.getDataRange().getValues();
  // 1행 타이틀, 2행 헤더, 3행부터 데이터
  const header = values[1];
  const ci = {
    send: idxOf_(header, '발송'), name: idxOf_(header, '성명'), email: idxOf_(header, '이메일'),
    bmonth: idxOf_(header, '생일월'), status: idxOf_(header, '상태'), empno: idxOf_(header, '사원번호'),
    subject: idxOf_(header, '제목'), body: idxOf_(header, '본문(실제 발송 내용)')
  };
  if (ci.send < 0 || ci.email < 0 || ci.body < 0) { ui.alert('검토 시트 형식이 올바르지 않습니다. ①을 다시 실행하세요.'); return; }

  const year = new Date().getFullYear();
  const sentKeys = CONFIG.RESEND ? {} : loadSentKeys_();
  const targets = [];
  for (let r = 2; r < values.length; r++) {
    const row = values[r];
    if (row[ci.send] !== true) continue;                 // 체크 안 됨
    const email = String(row[ci.email] || '').trim();
    if (!isValidEmail_(email)) continue;                 // 이메일 없음/이상
    const empno = String(row[ci.empno] || '').trim() || email;
    const month = parseInt(row[ci.bmonth], 10) || (new Date().getMonth() + 1);
    const key = empno + '|' + year + ('0' + month).slice(-2);
    if (sentKeys[key]) continue;                          // 이미 발송(중복 방지)
    targets.push({ rowNum: r + 1, email, name: String(row[ci.name] || ''), key,
                   subject: String(row[ci.subject] || ''), body: String(row[ci.body] || '') });
  }

  if (!targets.length) { ui.alert('발송할 대상이 없습니다. (체크됨/유효 이메일/미발송 조건을 확인하세요)'); return; }
  const quota = MailApp.getRemainingDailyQuota();
  if (quota < targets.length) { ui.alert('일일 한도 부족: 대상 ' + targets.length + ' / 남은 ' + quota + '통'); return; }

  if (ui.alert('발송 확인', '체크된 ' + targets.length + '명에게 검토 시트의 본문 그대로 발송합니다.\n계속할까요?',
      ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  const log = getLog_();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  let ok = 0, fail = 0; const fails = [];
  targets.forEach(t => {
    try {
      GmailApp.sendEmail(t.email, t.subject, t.body, { htmlBody: textToHtml_(t.body), name: CONFIG.SENDER_NAME });
      log.appendRow([t.key, '', t.name, t.email, stamp]);
      sh.getRange(t.rowNum, ci.status + 1).setValue('발송완료');
      sh.getRange(t.rowNum, 1, 1, HEADERS.length).setBackground('#e7f4e4');
      ok++;
    } catch (e) { fail++; fails.push(t.name + ': ' + e.message); }
    Utilities.sleep(300);
  });

  let msg = '발송 완료\n성공 ' + ok + '명 · 실패 ' + fail + '명\n남은 일일 한도 약 ' + MailApp.getRemainingDailyQuota() + '통';
  if (fails.length) msg += '\n\n[실패]\n' + fails.join('\n');
  ui.alert(msg);
}


// ===== 테스트 / 관리 =====
function sendTestToMe() {
  const ui = SpreadsheetApp.getUi();
  const me = Session.getActiveUser().getEmail();
  if (!me) { ui.alert('본인 이메일을 확인할 수 없습니다.'); return; }
  const month = new Date().getMonth() + 1;
  const m = buildMessage_('테스트', month);
  GmailApp.sendEmail(me, '[테스트] ' + m.subject, m.text, { htmlBody: textToHtml_(m.text), name: CONFIG.SENDER_NAME });
  ui.alert('테스트 메일을 ' + me + ' 로 보냈습니다.');
}

function logKey_(empno, year, month) { return empno + '|' + year + ('0' + month).slice(-2); }
function getLog_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) { sh = ss.insertSheet(LOG_SHEET); sh.appendRow(['키','사원번호','성명','이메일','발송일시']); sh.hideSheet(); }
  return sh;
}
function loadSentKeys_() {
  const v = getLog_().getDataRange().getValues(); const set = {};
  for (let i = 1; i < v.length; i++) set[String(v[i][0])] = true;
  return set;
}
function resetLog() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('발송기록 초기화', LOG_SHEET + ' 기록을 모두 지웁니다(중복방지 초기화). 계속할까요?', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const sh = getLog_(); sh.clearContents(); sh.appendRow(['키','사원번호','성명','이메일','발송일시']);
  ui.alert('초기화했습니다.');
}
function showQuota() {
  SpreadsheetApp.getUi().alert('오늘 남은 Gmail 발송 한도: 약 ' + MailApp.getRemainingDailyQuota() +
    '통\n(개인 500 / 워크스페이스 1,500 기준, 다음 날 초기화)');
}
