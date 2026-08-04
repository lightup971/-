/**
 * 생일휴가 안내 – 사원명부(구글 스프레드시트) → Gmail 발송
 * ---------------------------------------------------------------
 * ① 대상 분류: 선택한 월의 생일자를 '대상 여부/제외 사유/발송 여부'에 표시
 * ② 나에게 테스트 발송
 * ③ '발송 여부'가 체크된 사람에게만 발송
 *
 * [이모지] HTML 본문에는 숫자 문자 참조(&#127874;)로 넣어 인코딩 문제로 깨지지 않습니다.
 * [서명]   Gmail 설정에 등록된 본인 서명(로고 포함)을 Gmail API로 읽어와 그대로 붙입니다.
 *          → Apps Script 편집기 왼쪽 '서비스(Services) +' → Gmail 선택 → 추가 (1회 설정)
 */

// ===== 설정 =====
var CONFIG = {
  SHEET_NAME: '',
  SENDER_NAME: '부릉 피플실',
  SUBJECT: '[부릉] {월}월 생일 축하와 생일휴가 안내',
  EMOJI_IN_SUBJECT: false,           // 제목 이모지: 깨짐 위험이 있어 기본 꺼둠(true로 켜면 케이크 이모지 추가)
  INCLUDE_STATUSES: ['재직'],
  USE_GIVEN_NAME_ONLY: true,
  RESEND: false,

  EXCLUDE_RANKS: ['CEO','CTO','대표이사','부사장','전무','상무','이사','LV.8'],
  EXCLUDE_DEPTS: ['장애인고용'],

  USE_GMAIL_SIGNATURE: true,         // Gmail 설정의 내 서명을 그대로 사용(로고 포함)
  FALLBACK_SIGNATURE_HTML:           // 위 서명을 못 읽을 때만 사용되는 대체 서명
    '<div style="color:#5f5e5a;">──────────────<br>주식회사 부릉(VROONG) People실</div>'
};

// 이모지 코드포인트 (HTML에서는 &#숫자; 로 출력 → 절대 깨지지 않음)
var EMO = { cake: 127874, party: 127881, cal: 128197, pin: 128204, mega: 128226 };

// 본문 템플릿 — 이모지는 [[CAKE]] 같은 토큰으로 두고, 발송 시 변환합니다.
var BODY_TEXT =
'안녕하세요, {이름}님.\n' +
'부릉 피플실입니다.\n' +
'\n' +
'{월}월 생일을 진심으로 축하드립니다! [[PARTY]]\n' +
'생일 당월을 더욱 행복하게 보내실 수 있도록,\n' +
'부릉에서는 생일휴가 1일을 드리고 있습니다.\n' +
'\n' +
'[[CAL]] 생일휴가 1일 부여\n' +
'- 신청·사용 기한: 생일 당월인 {월}월 말일({말일})까지\n' +
'- 말일 이후 자동 소멸되며, 당월 내에서만 사용 가능합니다.\n' +
'\n' +
'[[PIN]] 신청 방법\n' +
'- 옴니이솔 → 인사관리 → My HR(ESS) → 근태신청(NEW) → [휴가신청서]\n' +
'  → 휴가종류: 기타휴가 > 생일휴가 선택\n' +
'\n' +
'[[MEGA]] 결재 라인\n' +
'- 1차 조직장(결재) → 피플실(합의)\n' +
'- ※ 반드시 \'피플실(합의)\'을 지정해 주세요.\n' +
'\n' +
'뜻깊고 즐거운 생일 보내시길 바랍니다.\n' +
'감사합니다.';

var LOG_SHEET = '_발송기록';

var TOKENS = [
  { t: '[[CAKE]]',  c: EMO.cake },
  { t: '[[PARTY]]', c: EMO.party },
  { t: '[[CAL]]',   c: EMO.cal },
  { t: '[[PIN]]',   c: EMO.pin },
  { t: '[[MEGA]]',  c: EMO.mega }
];

// 토큰 → HTML 숫자 문자 참조(깨지지 않음)
function tokensToHtml_(s) {
  TOKENS.forEach(function (o) { s = s.split(o.t).join('&#' + o.c + ';'); });
  return s;
}
// 토큰 → 실제 이모지 문자(평문 대체본용)
function tokensToPlain_(s) {
  TOKENS.forEach(function (o) { s = s.split(o.t).join(String.fromCodePoint(o.c)); });
  return s;
}
// 토큰 제거(제목 등)
function tokensStrip_(s) {
  TOKENS.forEach(function (o) { s = s.split(o.t).join(''); });
  return s.replace(/\s+$/,'');
}


// ===== 메뉴 =====
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('생일휴가 안내')
    .addItem('① 대상 분류(월 선택)', 'markTargets')
    .addItem('② 나에게 테스트 발송', 'sendTestToMe')
    .addItem('③ 체크된 사람에게 발송', 'sendMarked')
    .addSeparator()
    .addItem('내 Gmail 서명 확인', 'checkSignature')
    .addItem('발송기록 초기화', 'resetLog')
    .addItem('남은 일일 발송량 확인', 'showQuota')
    .addToUi();
}


// ===== 시트/컬럼 =====
function getRosterSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (CONFIG.SHEET_NAME) {
    var s = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!s) throw new Error('명단 시트를 찾을 수 없습니다: ' + CONFIG.SHEET_NAME);
    return s;
  }
  var active = ss.getActiveSheet();
  if (active.getName() !== LOG_SHEET) return active;
  var others = ss.getSheets().filter(function (s) { return s.getName() !== LOG_SHEET; });
  if (!others.length) throw new Error('명단 시트를 찾을 수 없습니다.');
  return others[0];
}

function findCol_(headers, candidates, exactOnly) {
  var norm = headers.map(function (h) { return String(h).trim(); });
  for (var i = 0; i < candidates.length; i++) { var j = norm.indexOf(candidates[i]); if (j >= 0) return j; }
  if (!exactOnly) for (var k = 0; k < candidates.length; k++) {
    var cand = candidates[k];
    var idx = norm.findIndex(function (h) { return h.toLowerCase().indexOf(cand.toLowerCase()) >= 0; });
    if (idx >= 0) return idx;
  }
  return -1;
}

function parse_() {
  var sheet = getRosterSheet_();
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) throw new Error('명단 데이터가 없습니다. 1행 머리글 아래에 명단을 넣으세요.');
  var h = values[0];
  var col = {
    empno:  findCol_(h, ['사원번호','사번']),
    name:   findCol_(h, ['성명','이름','name']),
    birth:  findCol_(h, ['생년월일','생일','출생','birth','dob']),
    bmonth: findCol_(h, ['생일월'], true),
    email:  findCol_(h, ['이메일','메일','email','gmail']),
    status: findCol_(h, ['재직여부','재직상태','status']),
    dept:   findCol_(h, ['비용센터','부서']),
    rank:   findCol_(h, ['직급'], true),
    role:   findCol_(h, ['직책'], true),
    title:  findCol_(h, ['직위'], true),
    targetYN: findCol_(h, ['대상 여부','대상여부'], true),
    reason:   findCol_(h, ['제외 사유','제외사유'], true),
    sendFlag: findCol_(h, ['발송 여부','발송여부'], true)
  };
  if (col.email < 0) throw new Error("'이메일' 컬럼을 찾을 수 없습니다.");
  if (col.birth < 0 && col.bmonth < 0) throw new Error("'생년월일' 또는 '생일월' 컬럼이 필요합니다.");
  return { sheet: sheet, values: values, col: col };
}

function lastDayOfMonth_(y, m) { return new Date(y, m, 0).getDate(); }
function monthFromBirth_(v) {
  if (v == null || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return v.getMonth() + 1;
  var m = String(v).match(/^\s*\d{4}[-./](\d{1,2})/); if (m) return parseInt(m[1], 10);
  m = String(v).match(/^\s*(\d{1,2})\s*$/); return m ? parseInt(m[1], 10) : null;
}
function rowMonth_(row, col) {
  return col.birth >= 0 ? monthFromBirth_(row[col.birth]) : monthFromBirth_(row[col.bmonth]);
}
function isValidEmail_(e) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(e || '').trim()); }
function greetingName_(n) { n = String(n || '').trim(); return CONFIG.USE_GIVEN_NAME_ONLY && n.length > 1 ? n.slice(1) : n; }

function excludeReason_(row, col) {
  if (col.status >= 0) {
    var st = String(row[col.status] || '').trim();
    if (st && CONFIG.INCLUDE_STATUSES.indexOf(st) < 0) return st || '재직 아님';
  }
  var ranks = [col.rank, col.role, col.title];
  for (var i = 0; i < ranks.length; i++) if (ranks[i] >= 0) {
    var v = String(row[ranks[i]] || '').trim();
    if (v && CONFIG.EXCLUDE_RANKS.indexOf(v) >= 0) return '임원 제외';
  }
  if (col.dept >= 0) {
    var d = String(row[col.dept] || '').trim();
    if (CONFIG.EXCLUDE_DEPTS.indexOf(d) >= 0) return '부서 제외(' + d + ')';
  }
  return null;
}

// 제목/본문(토큰 포함) 생성
function buildMessage_(fullName, month) {
  var y = new Date().getFullYear();
  var lastLabel = month + '월 ' + lastDayOfMonth_(y, month) + '일';
  var body = String(BODY_TEXT)
    .replace(/{이름}/g, greetingName_(fullName))
    .replace(/{월}/g, month)
    .replace(/{말일}/g, lastLabel);
  var subject = String(CONFIG.SUBJECT).replace(/{월}/g, month);
  if (CONFIG.EMOJI_IN_SUBJECT) subject += ' ' + String.fromCodePoint(EMO.cake);
  return { subject: subject, body: body };
}

// 본문(토큰) → HTML
function bodyToHtml_(body) {
  var esc = function (s) { return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); };
  var headTokens = ['[[CAL]]','[[PIN]]','[[MEGA]]'];
  var html = String(body).split('\n').map(function (line) {
    var t = line.trim();
    var isHead = false;
    for (var i = 0; i < headTokens.length; i++) if (t.indexOf(headTokens[i]) === 0) isHead = true;
    if (isHead) return '<div style="margin:16px 0 6px;font-weight:600;">' + esc(line) + '</div>';
    if (t === '') return '<div style="height:8px;"></div>';
    if (t.indexOf('- ') === 0) return '<div style="padding-left:14px;">&#8226; ' + esc(t.slice(2)) + '</div>';
    return '<div>' + esc(line) + '</div>';
  }).join('');
  html = tokensToHtml_(html);   // 이모지 토큰 → &#숫자; (깨지지 않음)
  return '<div style="font-family:\'Apple SD Gothic Neo\',\'Malgun Gothic\',sans-serif;font-size:15px;line-height:1.7;color:#1a1a18;max-width:600px;">' + html + '</div>';
}

// Gmail 설정에 등록된 내 서명(HTML) 가져오기 — Gmail 고급 서비스 필요
function gmailSignatureHtml_() {
  if (!CONFIG.USE_GMAIL_SIGNATURE) return '';
  try {
    var me = Session.getActiveUser().getEmail();
    var res = Gmail.Users.Settings.SendAs.list('me');
    var list = (res && res.sendAs) ? res.sendAs : [];
    var i;
    for (i = 0; i < list.length; i++) {
      if (list[i].sendAsEmail === me && list[i].signature) return list[i].signature;
    }
    for (i = 0; i < list.length; i++) {
      if (list[i].isDefault && list[i].signature) return list[i].signature;
    }
    for (i = 0; i < list.length; i++) {
      if (list[i].signature) return list[i].signature;
    }
  } catch (e) {
    // Gmail 고급 서비스 미설정 등 → 대체 서명 사용
  }
  return '';
}

function signatureHtml_() {
  var sig = gmailSignatureHtml_();
  if (sig) return '<br><br>' + sig;
  return CONFIG.FALLBACK_SIGNATURE_HTML ? '<br><br>' + CONFIG.FALLBACK_SIGNATURE_HTML : '';
}

// 실제 발송
function sendMail_(email, subject, body) {
  var html = bodyToHtml_(body) + signatureHtml_();
  GmailApp.sendEmail(email, subject, tokensToPlain_(body), {
    htmlBody: html,
    name: CONFIG.SENDER_NAME
  });
}


// ===== ① 대상 분류 =====
function markTargets() {
  var ui = SpreadsheetApp.getUi();
  var p = parse_();
  var col = p.col, values = p.values, sheet = p.sheet;
  if (col.targetYN < 0 || col.reason < 0 || col.sendFlag < 0) {
    ui.alert("명단 시트에 '대상 여부', '제외 사유', '발송 여부' 컬럼이 필요합니다."); return;
  }

  var nowMonth = new Date().getMonth() + 1;
  var resp = ui.prompt('대상 분류', '대상 월(1~12)을 입력하세요.\n기본값: ' + nowMonth + '월', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var input = resp.getResponseText().trim();
  var month = input ? parseInt(input, 10) : nowMonth;
  if (!(month >= 1 && month <= 12)) { ui.alert('1~12 사이로 입력하세요.'); return; }

  var n = values.length - 1;
  var yn = [], rs = [], sf = [];
  var cT = 0, cE = 0, cN = 0;
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (col.name >= 0 && !String(row[col.name] || '').trim()) { yn.push(['']); rs.push(['']); sf.push([false]); continue; }
    if (rowMonth_(row, col) !== month) { yn.push(['']); rs.push(['']); sf.push([false]); continue; }
    var reason = excludeReason_(row, col);
    var email = String(row[col.email] || '').trim();
    if (reason) { yn.push(['제외']); rs.push([reason]); sf.push([false]); cE++; }
    else if (!isValidEmail_(email)) { yn.push(['이메일없음']); rs.push(['']); sf.push([false]); cN++; }
    else { yn.push(['대상']); rs.push(['']); sf.push([true]); cT++; }
  }

  sheet.getRange(2, col.targetYN + 1, n, 1).setValues(yn);
  sheet.getRange(2, col.reason + 1, n, 1).setValues(rs);
  var sfRange = sheet.getRange(2, col.sendFlag + 1, n, 1);
  sfRange.insertCheckboxes();
  sfRange.setValues(sf);

  ui.alert(month + '월 분류 완료\n- 대상 ' + cT + '명(발송 여부 체크됨) · 제외 ' + cE + '명 · 이메일없음 ' + cN + '명\n\n' +
    "'발송 여부' 열을 확인/수정한 뒤, 메뉴 ③으로 발송하세요.");
}


// ===== ③ 발송 =====
function sendMarked() {
  var ui = SpreadsheetApp.getUi();
  var p = parse_();
  var col = p.col, values = p.values, sheet = p.sheet;
  if (col.sendFlag < 0) { ui.alert("'발송 여부' 컬럼을 찾을 수 없습니다. 먼저 ①을 실행하세요."); return; }

  var year = new Date().getFullYear();
  var sentKeys = CONFIG.RESEND ? {} : loadSentKeys_();
  var targets = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var flag = row[col.sendFlag];
    var checked = (flag === true) || (String(flag).toUpperCase() === 'TRUE');
    if (!checked) continue;
    var email = String(row[col.email] || '').trim();
    if (!isValidEmail_(email)) continue;
    var name = col.name >= 0 ? String(row[col.name]).trim() : '';
    var empno = col.empno >= 0 ? String(row[col.empno] || '').trim() : email;
    var month = rowMonth_(row, col) || (new Date().getMonth() + 1);
    var key = empno + '|' + year + ('0' + month).slice(-2);
    if (sentKeys[key]) continue;
    targets.push({ rowNum: r + 1, email: email, name: name, empno: empno, month: month, key: key });
  }

  if (!targets.length) { ui.alert('발송할 대상이 없습니다. (발송 여부 체크 / 유효 이메일 / 미발송 조건 확인)'); return; }
  var quota = MailApp.getRemainingDailyQuota();
  if (quota < targets.length) { ui.alert('일일 한도 부족: 대상 ' + targets.length + ' / 남은 ' + quota + '통'); return; }

  if (ui.alert('발송 확인', '체크된 ' + targets.length + '명에게 생일휴가 안내를 발송합니다.\n계속할까요?',
      ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  var log = getLog_();
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var ok = 0, fail = 0, fails = [];
  targets.forEach(function (t) {
    try {
      var m = buildMessage_(t.name, t.month);
      sendMail_(t.email, m.subject, m.body);
      log.appendRow([t.key, t.empno, t.name, t.email, stamp]);
      if (col.targetYN >= 0) sheet.getRange(t.rowNum, col.targetYN + 1).setValue('발송완료');
      sheet.getRange(t.rowNum, col.sendFlag + 1).setValue(false);
      ok++;
    } catch (e) { fail++; fails.push(t.name + ': ' + e.message); }
    Utilities.sleep(300);
  });

  var msg = '발송 완료\n성공 ' + ok + '명 · 실패 ' + fail + '명\n남은 일일 한도 약 ' + MailApp.getRemainingDailyQuota() + '통';
  if (fails.length) msg += '\n\n[실패]\n' + fails.join('\n');
  ui.alert(msg);
}


// ===== 테스트 / 관리 =====
function sendTestToMe() {
  var ui = SpreadsheetApp.getUi();
  var me = Session.getActiveUser().getEmail();
  if (!me) { ui.alert('본인 이메일을 확인할 수 없습니다.'); return; }
  var month = new Date().getMonth() + 1;
  var m = buildMessage_('홍길동', month);
  sendMail_(me, '[테스트] ' + m.subject, m.body);
  var sigOn = gmailSignatureHtml_() ? '내 Gmail 서명 적용됨' : '대체 서명 사용(내 Gmail 서명 못 읽음)';
  ui.alert('테스트 메일을 ' + me + ' 로 보냈습니다.\n(예시 이름: 홍길동 · ' + sigOn + ')');
}

function checkSignature() {
  var sig = gmailSignatureHtml_();
  if (sig) {
    SpreadsheetApp.getUi().alert('내 Gmail 서명을 정상적으로 읽었습니다.\n메일 하단에 이 서명이 그대로 들어갑니다.\n\n(길이 ' + sig.length + '자)');
  } else {
    SpreadsheetApp.getUi().alert(
      '내 Gmail 서명을 읽지 못했습니다.\n\n' +
      'Apps Script 편집기 왼쪽 [서비스(Services)] 옆 + 클릭 → 목록에서 Gmail 선택 → 추가 후\n' +
      '다시 시도하세요. (그전까지는 대체 서명이 사용됩니다.)');
  }
}

function getLog_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) { sh = ss.insertSheet(LOG_SHEET); sh.appendRow(['키','사원번호','성명','이메일','발송일시']); sh.hideSheet(); }
  return sh;
}
function loadSentKeys_() {
  var v = getLog_().getDataRange().getValues(); var set = {};
  for (var i = 1; i < v.length; i++) set[String(v[i][0])] = true;
  return set;
}
function resetLog() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('발송기록 초기화', LOG_SHEET + ' 기록을 모두 지웁니다(중복방지 초기화). 계속할까요?', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  var sh = getLog_(); sh.clearContents(); sh.appendRow(['키','사원번호','성명','이메일','발송일시']);
  ui.alert('초기화했습니다.');
}
function showQuota() {
  SpreadsheetApp.getUi().alert('오늘 남은 Gmail 발송 한도: 약 ' + MailApp.getRemainingDailyQuota() +
    '통\n(개인 500 / 워크스페이스 1,500 기준, 다음 날 초기화)');
}
