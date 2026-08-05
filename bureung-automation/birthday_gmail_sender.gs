/**
 * 생일휴가 안내 – 사원명부(구글 스프레드시트) → Gmail 발송
 * ===============================================================
 * 유지보수 원칙: 담당자가 바뀌어도 '스프레드시트만' 보면 되도록 설계.
 *   - 메일 제목/본문, 제외 규칙, 발신자명 → [설정] 시트에서 수정 (코드 수정 불필요)
 *   - 사용법/인수인계 절차          → [사용법] 시트에 내장
 *   - 처음 열었다면 메뉴의 '설정·사용법 시트 만들기'를 먼저 실행하세요.
 */

var SETTINGS_SHEET = '설정';
var GUIDE_SHEET = '사용법';
var LOG_SHEET = '_발송기록';

// [설정] 시트가 없을 때 사용되는 기본값 (설정 시트가 있으면 그쪽이 우선)
var DEFAULTS = {
  senderName: '부릉 피플실',
  subject: '[피플실] {연도}년 {월}월 생일휴가 부여 및 사용 안내(~{말일짧게}까지★)',
  statuses: '재직',
  excludeRanks: 'CEO,CTO,대표이사,부사장,전무,상무,이사,LV.8',
  excludeDepts: '장애인고용',
  nameStyle: '성 제외',
  body: [
    '안녕하세요, {이름}님.',
    '부릉 피플실입니다.',
    '',
    '{월}월 생일을 진심으로 축하드립니다! [[PARTY]]',
    '',
    '생일 당월을 더욱 행복하게 보내실 수 있도록, 부릉에서는 생일휴가 1일을 드리고 있습니다.',
    '',
    '[[CAL]] 생일휴가 1일 부여',
    '- 신청·사용 기한: 생일 당월인 {월}월 말일({말일})까지',
    '- 말일 이후 자동 소멸되며, 당월 내에서만 사용 가능합니다.',
    '',
    '[[PIN]] 신청 방법',
    '- 옴니이솔 → 인사관리 → My HR(ESS) → 근태신청(NEW) → [휴가신청서]',
    '- 휴가종류: 기타휴가 > 생일휴가 선택',
    '',
    '[[MEGA]] 결재 라인',
    '- 1차 조직장(결재) → 피플실(합의)',
    '- ※ 반드시 \'피플실(합의)\'을 지정해 주세요.',
    '',
    '',
    '뜻깊고 즐거운 생일 보내시길 바랍니다.',
    '',
    '감사합니다.'
  ].join('\n')
};

// 코드에 실제 이모지를 쓰면 붙여넣기 과정에서 깨질 수 있어 토큰으로 둡니다.
// (설정 시트에 직접 입력하는 이모지는 안전하므로 그대로 쓰셔도 됩니다.)
var TOKENS = [
  { t: '[[CAKE]]',  c: 127874 },
  { t: '[[PARTY]]', c: 127881 },
  { t: '[[CAL]]',   c: 128197 },
  { t: '[[PIN]]',   c: 128204 },
  { t: '[[MEGA]]',  c: 128226 }
];
function tokensToPlain_(s) {
  TOKENS.forEach(function (o) { s = s.split(o.t).join(String.fromCodePoint(o.c)); });
  return s;
}
// 이모지(4바이트 문자)를 HTML 숫자 참조로 바꿔 인코딩 문제로 깨지지 않게 합니다.
function emojiToEntities_(s) {
  var out = '';
  for (var i = 0; i < s.length; i++) {
    var cp = s.codePointAt(i);
    if (cp > 0xFFFF) { out += '&#' + cp + ';'; i++; }
    else out += s.charAt(i);
  }
  return out;
}
function startsWithEmoji_(s) {
  var t = s.replace(/^\s+/, '');
  return t.length > 0 && t.codePointAt(0) > 0xFFFF;
}


// ===== 메뉴 =====
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('생일휴가 안내')
    .addItem('① 대상 분류(월 선택)', 'markTargets')
    .addItem('② 나에게 테스트 발송', 'sendTestToMe')
    .addItem('③ 체크된 사람에게 발송', 'sendMarked')
    .addSeparator()
    .addItem('설정·사용법 시트 만들기', 'setupSheets')
    .addItem('내 Gmail 서명 확인', 'checkSignature')
    .addSeparator()
    .addItem('발송기록 초기화', 'resetLog')
    .addItem('남은 일일 발송량 확인', 'showQuota')
    .addToUi();
}


// ===== 설정 =====
function settingsRows_() {
  return [
    ['발신자 이름',        DEFAULTS.senderName,   '받는 사람에게 보이는 이름'],
    ['메일 제목',          DEFAULTS.subject,      '{연도} {월} {말일짧게}(예: 8/31) 자동 치환'],
    ['발송 대상 재직상태', DEFAULTS.statuses,     '쉼표로 여러 개 가능 (예: 재직,휴직)'],
    ['제외 직급/직책',     DEFAULTS.excludeRanks, '이 직급·직책이면 발송 대상에서 제외 (쉼표 구분)'],
    ['제외 부서',          DEFAULTS.excludeDepts, '이 부서면 발송 대상에서 제외 (쉼표 구분)'],
    ['이름 표기',          DEFAULTS.nameStyle,    "'성 제외' = 김민희→민희님 / '전체 이름' = 김민희님"],
    ['메일 본문',          tokensToPlain_(DEFAULTS.body),
     '{이름} {월} {말일} 자동 치환. 이모지로 시작하는 줄은 소제목으로 굵게 표시됩니다.']
  ];
}

function setupSheets() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (ss.getSheetByName(SETTINGS_SHEET)) {
    if (ui.alert('설정 시트 다시 만들기',
        "'" + SETTINGS_SHEET + "' 시트가 이미 있습니다.\n기본값으로 덮어쓸까요? (수정한 문구가 사라집니다)",
        ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
    ss.deleteSheet(ss.getSheetByName(SETTINGS_SHEET));
  }
  var sh = ss.insertSheet(SETTINGS_SHEET);
  sh.getRange(1, 1).setValue('생일휴가 안내 – 설정  (이 시트에서 문구와 규칙을 수정하세요. 코드는 건드릴 필요 없습니다.)');
  sh.getRange(1, 1, 1, 3).merge().setFontWeight('bold').setBackground('#fff3cd').setWrap(true);
  sh.getRange(2, 1, 1, 3).setValues([['항목', '값', '설명']]).setFontWeight('bold').setBackground('#eeeeee');
  var rows = settingsRows_();
  sh.getRange(3, 1, rows.length, 3).setValues(rows);
  sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 520); sh.setColumnWidth(3, 360);
  sh.getRange(3, 2, rows.length, 1).setWrap(true).setVerticalAlignment('top');
  sh.getRange(3, 3, rows.length, 1).setWrap(true).setVerticalAlignment('top').setFontColor('#666666');
  sh.setFrozenRows(2);

  buildGuideSheet_(ss);
  ui.alert("'" + SETTINGS_SHEET + "' 와 '" + GUIDE_SHEET + "' 시트를 만들었습니다.\n\n" +
           '앞으로 문구·규칙 수정은 [설정] 시트에서 하시면 됩니다.');
}

function readSettings_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SETTINGS_SHEET);
  var map = {};
  if (sh) {
    var v = sh.getDataRange().getValues();
    for (var r = 0; r < v.length; r++) {
      var k = String(v[r][0] || '').trim();
      if (k) map[k] = v[r][1];
    }
  }
  var get = function (key, dflt) {
    var x = map[key];
    return (x === undefined || String(x).trim() === '') ? dflt : String(x);
  };
  var list = function (s) {
    return String(s).split(',').map(function (x) { return x.trim(); }).filter(String);
  };
  return {
    senderName:   get('발신자 이름', DEFAULTS.senderName),
    subject:      get('메일 제목', DEFAULTS.subject),
    statuses:     list(get('발송 대상 재직상태', DEFAULTS.statuses)),
    excludeRanks: list(get('제외 직급/직책', DEFAULTS.excludeRanks)),
    excludeDepts: list(get('제외 부서', DEFAULTS.excludeDepts)),
    givenNameOnly: get('이름 표기', DEFAULTS.nameStyle).indexOf('전체') < 0,
    body:         tokensToPlain_(get('메일 본문', DEFAULTS.body))
  };
}


// ===== 명단 시트 =====
function getRosterSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var skip = [SETTINGS_SHEET, GUIDE_SHEET, LOG_SHEET];
  var active = ss.getActiveSheet();
  if (skip.indexOf(active.getName()) < 0) return active;
  var others = ss.getSheets().filter(function (s) { return skip.indexOf(s.getName()) < 0; });
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

// 머리글이 1행이 아닐 수 있으므로(제목/메모 행이 위에 있는 경우) 자동으로 찾습니다.
function detectHeaderRow_(values) {
  var known = ['사원번호','사번','성명','이름','비용센터','부서','직급','직책','직위',
               '재직여부','생년월일','이메일','메일','생일월','대상 여부','제외 사유','발송 여부'];
  var bestRow = 0, bestScore = -1;
  var limit = Math.min(values.length, 20);
  for (var r = 0; r < limit; r++) {
    var row = values[r], score = 0;
    for (var c = 0; c < row.length; c++) {
      var v = String(row[c] || '').trim();
      if (v && known.indexOf(v) >= 0) score++;
    }
    if (score > bestScore) { bestScore = score; bestRow = r; }
  }
  return bestRow;
}

function parse_() {
  var sheet = getRosterSheet_();
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) throw new Error('명단 데이터가 없습니다. 머리글 아래에 명단을 넣으세요.');
  var headerRow = detectHeaderRow_(values);
  var h = values[headerRow];
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
  if (col.email < 0) {
    throw new Error("'이메일' 컬럼을 찾을 수 없습니다.\n인식된 머리글 행: " + (headerRow + 1) + '행\n' +
      '내용: ' + h.map(function (x) { return String(x || ''); }).filter(String).join(' | '));
  }
  if (col.birth < 0 && col.bmonth < 0) throw new Error("'생년월일' 또는 '생일월' 컬럼이 필요합니다.");
  return { sheet: sheet, values: values, col: col, headerRow: headerRow };
}

function lastDayOfMonth_(y, m) { return new Date(y, m, 0).getDate(); }
function monthFromBirth_(v) {
  if (v == null || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return v.getMonth() + 1;
  var m = String(v).match(/^\s*\d{4}[-./](\d{1,2})/); if (m) return parseInt(m[1], 10);
  m = String(v).match(/^\s*(\d{1,2})\s*$/); return m ? parseInt(m[1], 10) : null;
}
// 생년월일에서 먼저 계산하고, 비어 있으면 '생일월' 값으로 대체
function rowMonth_(row, col) {
  var m = col.birth >= 0 ? monthFromBirth_(row[col.birth]) : null;
  if (m == null && col.bmonth >= 0) m = monthFromBirth_(row[col.bmonth]);
  return m;
}
function isValidEmail_(e) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(e || '').trim()); }
function greetingName_(n, cfg) {
  n = String(n || '').trim();
  return (cfg.givenNameOnly && n.length > 1) ? n.slice(1) : n;
}

function excludeReason_(row, col, cfg) {
  if (col.status >= 0) {
    var st = String(row[col.status] || '').trim();
    if (st && cfg.statuses.indexOf(st) < 0) return st || '재직 아님';
  }
  var ranks = [col.rank, col.role, col.title];
  for (var i = 0; i < ranks.length; i++) if (ranks[i] >= 0) {
    var v = String(row[ranks[i]] || '').trim();
    if (v && cfg.excludeRanks.indexOf(v) >= 0) return '임원 제외';
  }
  if (col.dept >= 0) {
    var d = String(row[col.dept] || '').trim();
    if (cfg.excludeDepts.indexOf(d) >= 0) return '부서 제외(' + d + ')';
  }
  return null;
}


// ===== 메일 =====
function buildMessage_(fullName, month, cfg) {
  var y = new Date().getFullYear();
  var last = lastDayOfMonth_(y, month);
  var rep = function (s) {
    return String(s).replace(/{이름}/g, greetingName_(fullName, cfg))
                    .replace(/{연도}/g, y)
                    .replace(/{월}/g, month)
                    .replace(/{말일짧게}/g, month + '/' + last)   // 8/31
                    .replace(/{말일}/g, month + '월 ' + last + '일'); // 8월 31일
  };
  return { subject: rep(cfg.subject), body: rep(cfg.body) };
}

function bodyToHtml_(body) {
  var esc = function (s) { return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); };
  var html = String(body).split('\n').map(function (line) {
    var t = line.trim();
    if (t === '') return '<div style="height:8px;"></div>';
    // 이모지로 시작하는 줄 = 소제목
    if (startsWithEmoji_(t)) return '<div style="margin:16px 0 6px;font-weight:600;">' + esc(line) + '</div>';
    if (t.indexOf('- ') === 0) return '<div style="padding-left:26px;text-indent:-13px;">&#8226; ' + esc(t.slice(2)) + '</div>';
    return '<div>' + esc(line) + '</div>';
  }).join('');
  html = emojiToEntities_(html);
  // word-break:keep-all → 한글이 단어 중간에서 잘리지 않고 어절 단위로 줄바꿈
  return '<div style="font-family:\'Apple SD Gothic Neo\',\'Malgun Gothic\',sans-serif;' +
         'font-size:15px;line-height:1.7;color:#1a1a18;max-width:760px;' +
         'word-break:keep-all;overflow-wrap:break-word;">' + html + '</div>';
}

// Gmail 설정에 등록된 내 서명(로고 포함)을 가져옵니다 — Gmail 고급 서비스 필요
function gmailSignatureResult_() {
  if (typeof Gmail === 'undefined') {
    return { ok: false, why: 'Gmail 고급 서비스가 추가되지 않았습니다.\n편집기 왼쪽 [서비스] 옆 + → Gmail → 추가' };
  }
  try {
    var me = Session.getActiveUser().getEmail();
    var res = Gmail.Users.Settings.SendAs.list('me');
    var list = (res && res.sendAs) ? res.sendAs : [];
    if (!list.length) return { ok: false, why: '보내는 주소(SendAs) 정보를 가져오지 못했습니다.' };
    var i, pick = '';
    for (i = 0; i < list.length; i++) if (list[i].sendAsEmail === me && list[i].signature) { pick = list[i].signature; break; }
    if (!pick) for (i = 0; i < list.length; i++) if (list[i].isDefault && list[i].signature) { pick = list[i].signature; break; }
    if (!pick) for (i = 0; i < list.length; i++) if (list[i].signature) { pick = list[i].signature; break; }
    if (!pick) {
      return { ok: false, why: 'Gmail에서 서명을 찾지 못했습니다.\n조회된 주소: ' +
               list.map(function (x) { return x.sendAsEmail; }).join(', ') };
    }
    return { ok: true, signature: pick };
  } catch (e) {
    return { ok: false, why: '오류: ' + e.message +
      '\n\n서비스 추가 후에는 권한 재승인이 필요합니다. 편집기에서 함수를 한 번 실행해 승인하세요.' };
  }
}

function sendMail_(email, subject, body, cfg) {
  var sig = gmailSignatureResult_();
  var html = bodyToHtml_(body) + (sig.ok ? '<br><br>' + sig.signature : '');
  GmailApp.sendEmail(email, subject, body, { htmlBody: html, name: cfg.senderName });
}


// ===== ① 대상 분류 =====
function markTargets() {
  var ui = SpreadsheetApp.getUi();
  var cfg = readSettings_();
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

  var first = p.headerRow + 1;
  var n = values.length - first;
  if (n < 1) { ui.alert('명단 데이터가 없습니다.'); return; }
  var yn = [], rs = [], sf = [];
  var cT = 0, cE = 0, cN = 0;
  for (var r = first; r < values.length; r++) {
    var row = values[r];
    if (col.name >= 0 && !String(row[col.name] || '').trim()) { yn.push(['']); rs.push(['']); sf.push([false]); continue; }
    if (rowMonth_(row, col) !== month) { yn.push(['']); rs.push(['']); sf.push([false]); continue; }
    var reason = excludeReason_(row, col, cfg);
    var email = String(row[col.email] || '').trim();
    if (reason) { yn.push(['제외']); rs.push([reason]); sf.push([false]); cE++; }
    else if (!isValidEmail_(email)) { yn.push(['이메일없음']); rs.push(['']); sf.push([false]); cN++; }
    else { yn.push(['대상']); rs.push(['']); sf.push([true]); cT++; }
  }

  sheet.getRange(first + 1, col.targetYN + 1, n, 1).setValues(yn);
  sheet.getRange(first + 1, col.reason + 1, n, 1).setValues(rs);
  var sfRange = sheet.getRange(first + 1, col.sendFlag + 1, n, 1);
  sfRange.insertCheckboxes();
  sfRange.setValues(sf);

  ui.alert(month + '월 분류 완료\n- 대상 ' + cT + '명(발송 여부 체크됨) · 제외 ' + cE + '명 · 이메일없음 ' + cN + '명\n\n' +
    "'발송 여부' 열을 확인/수정한 뒤, 메뉴 ③으로 발송하세요.");
}


// ===== ③ 발송 =====
function sendMarked() {
  var ui = SpreadsheetApp.getUi();
  var cfg = readSettings_();
  var p = parse_();
  var col = p.col, values = p.values, sheet = p.sheet;
  if (col.sendFlag < 0) { ui.alert("'발송 여부' 컬럼을 찾을 수 없습니다. 먼저 ①을 실행하세요."); return; }

  var year = new Date().getFullYear();
  var sentKeys = loadSentKeys_();
  var targets = [];
  var stat = { checked: 0, noEmail: 0, already: 0 };
  var alreadyNames = [];
  for (var r = p.headerRow + 1; r < values.length; r++) {
    var row = values[r];
    var flag = row[col.sendFlag];
    if (!(flag === true || String(flag).toUpperCase() === 'TRUE')) continue;
    stat.checked++;
    var email = String(row[col.email] || '').trim();
    var name = col.name >= 0 ? String(row[col.name]).trim() : '';
    if (!isValidEmail_(email)) { stat.noEmail++; continue; }
    var empno = col.empno >= 0 ? String(row[col.empno] || '').trim() : email;
    if (!empno) empno = email;
    var month = rowMonth_(row, col) || (new Date().getMonth() + 1);
    var key = empno + '|' + year + ('0' + month).slice(-2);
    if (sentKeys[key]) { stat.already++; alreadyNames.push(name || email); continue; }
    targets.push({ rowNum: r + 1, email: email, name: name, empno: empno, month: month, key: key });
  }

  if (!targets.length) {
    var why = '발송할 대상이 없습니다.\n\n';
    why += "· '발송 여부'가 체크된 행: " + stat.checked + '명\n';
    if (stat.noEmail) why += '· 이메일이 없거나 형식이 잘못되어 제외: ' + stat.noEmail + '명\n';
    if (stat.already) {
      why += '· 이번 달에 이미 발송되어 제외(중복 방지): ' + stat.already + '명\n';
      why += '   → ' + alreadyNames.slice(0, 10).join(', ') + (alreadyNames.length > 10 ? ' 외' : '') + '\n';
      why += "\n같은 사람에게 다시 보내려면 메뉴 [발송기록 초기화]를 실행한 뒤 다시 시도하세요.";
    }
    if (!stat.checked) why += "\n먼저 [① 대상 분류]를 실행하거나, '발송 여부' 열을 직접 체크하세요.";
    ui.alert(why);
    return;
  }
  var quota = MailApp.getRemainingDailyQuota();
  if (quota < targets.length) { ui.alert('일일 한도 부족: 대상 ' + targets.length + ' / 남은 ' + quota + '통'); return; }

  var confirmMsg = targets.length + '명에게 생일휴가 안내를 발송합니다.\n';
  if (stat.already) confirmMsg += '(이미 발송된 ' + stat.already + '명은 제외됩니다)\n';
  if (stat.noEmail) confirmMsg += '(이메일 없는 ' + stat.noEmail + '명은 제외됩니다)\n';
  confirmMsg += '\n계속할까요?';
  if (ui.alert('발송 확인', confirmMsg, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  var log = getLog_();
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var ok = 0, fail = 0, fails = [];
  targets.forEach(function (t) {
    try {
      var m = buildMessage_(t.name, t.month, cfg);
      sendMail_(t.email, m.subject, m.body, cfg);
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
  var cfg = readSettings_();
  var me = Session.getActiveUser().getEmail();
  if (!me) { ui.alert('본인 이메일을 확인할 수 없습니다.'); return; }
  var month = new Date().getMonth() + 1;
  var m = buildMessage_('홍길동', month, cfg);
  sendMail_(me, '[테스트] ' + m.subject, m.body, cfg);
  var sig = gmailSignatureResult_();
  ui.alert('테스트 메일을 ' + me + ' 로 보냈습니다.\n(예시 이름: 홍길동)\n서명: ' +
           (sig.ok ? '내 Gmail 서명 적용됨' : '적용 안 됨 — ' + sig.why));
}

function checkSignature() {
  var r = gmailSignatureResult_();
  if (r.ok) {
    var plain = String(r.signature).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
    SpreadsheetApp.getUi().alert('내 Gmail 서명을 정상적으로 읽었습니다.\n\n' +
      '길이: ' + r.signature.length + '자\n이미지 포함: ' + (/<img/i.test(r.signature) ? '예(로고 포함)' : '아니오') +
      '\n\n미리보기: ' + plain.slice(0, 200));
  } else {
    SpreadsheetApp.getUi().alert('내 Gmail 서명을 읽지 못했습니다.\n\n' + r.why);
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
  if (ui.alert('발송기록 초기화', LOG_SHEET + ' 기록을 모두 지웁니다(중복방지 초기화). 계속할까요?',
      ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  var sh = getLog_(); sh.clearContents(); sh.appendRow(['키','사원번호','성명','이메일','발송일시']);
  ui.alert('초기화했습니다.');
}
function showQuota() {
  SpreadsheetApp.getUi().alert('오늘 남은 Gmail 발송 한도: 약 ' + MailApp.getRemainingDailyQuota() +
    '통\n(개인 500 / 워크스페이스 1,500 기준, 다음 날 초기화)');
}


// ===== 사용법 / 인수인계 문서 =====
function buildGuideSheet_(ss) {
  var old = ss.getSheetByName(GUIDE_SHEET);
  if (old) ss.deleteSheet(old);
  var sh = ss.insertSheet(GUIDE_SHEET, 0);

  var H = '■';
  var lines = [
    ['생일휴가 안내 메일 자동화 – 사용법 및 인수인계 문서'],
    [''],
    [H + ' 이 도구가 하는 일'],
    ['매월 생일자를 명단에서 자동으로 골라, 생일휴가 안내 메일을 Gmail로 보냅니다.'],
    ['임원·제외 부서·퇴사/휴직자는 자동으로 걸러지고, 같은 달 중복발송도 자동으로 막습니다.'],
    [''],
    [H + ' 매월 사용 순서 (5분)'],
    ['1. 상단 메뉴 [생일휴가 안내] → [① 대상 분류(월 선택)] → 이번 달 입력'],
    ['   → 명단의 "대상 여부/제외 사유/발송 여부" 열이 자동으로 채워집니다.'],
    ['2. "발송 여부" 체크박스를 눈으로 확인합니다. (빼고 싶으면 체크 해제, 추가하려면 체크)'],
    ['3. [② 나에게 테스트 발송] → 내 메일함에서 문구·서명을 확인합니다.'],
    ['4. [③ 체크된 사람에게 발송] → 인원 확인 후 발송. 발송된 행은 "발송완료"로 바뀝니다.'],
    [''],
    [H + ' 문구·규칙을 바꾸고 싶을 때'],
    ['[설정] 시트에서 수정하세요. 코드(Apps Script)는 열 필요 없습니다.'],
    ['  - 메일 제목 / 메일 본문 : {이름} {월} {말일} 은 자동으로 채워집니다.'],
    ['  - 제외 직급/직책, 제외 부서 : 쉼표로 구분해 입력'],
    ['  - 발송 대상 재직상태 : 기본 "재직". 휴직자도 보내려면 "재직,휴직"'],
    ['  - 이름 표기 : "성 제외"(민희님) 또는 "전체 이름"(김민희님)'],
    [''],
    [H + ' 명단을 갱신할 때 (신규 입사자 반영)'],
    ['ERP(옴니이솔) 사원명부를 새로 내려받아 명단 시트에 머리글째 붙여넣으면 됩니다.'],
    ['이메일·생년월일이 그 파일에 들어 있으므로, 사람별로 따로 찾을 필요가 없습니다.'],
    ['필요한 열: 성명 / 생년월일(또는 생일월) / 이메일 / 재직여부 / 직급 / 직책 / 비용센터(부서) / 사원번호'],
    ['"대상 여부", "제외 사유", "발송 여부" 열은 비워두면 ①번 메뉴가 채웁니다.'],
    [''],
    [H + ' ★ 담당자가 바뀔 때 반드시 할 일 (중요)'],
    ['이 자동화는 "실행하는 사람의 구글 계정"으로 메일을 보냅니다.'],
    ['따라서 담당자가 바뀌면 아래를 반드시 진행하세요. 안 하면 어느 날 갑자기 멈춥니다.'],
    ['1. 이 스프레드시트의 소유권을 후임자(또는 팀 공용 계정)에게 이전합니다.'],
    ['   [공유] → 후임자 추가 → 점 3개 메뉴 → "소유권 이전"'],
    ['2. 후임자가 스프레드시트를 열고 [확장 프로그램] → [Apps Script] 진입'],
    ['3. 편집기 왼쪽 [서비스] 옆 + → Gmail 선택 → 추가  (서명을 쓰려면 필수)'],
    ['4. 메뉴 [② 나에게 테스트 발송] 실행 → 권한 승인 창에서 "고급" → "이동" → "허용"'],
    ['5. 테스트 메일이 정상 도착하면 인수인계 완료입니다.'],
    ['※ 메일 하단 서명은 "실행하는 사람의 Gmail 서명"이 자동으로 붙습니다.'],
    ['   후임자는 본인 Gmail 서명만 등록해두면 됩니다. (Gmail 설정 → 서명)'],
    [''],
    [H + ' 문제가 생겼을 때'],
    ['· "이메일 컬럼을 찾을 수 없습니다" → 명단 시트의 머리글(성명/이메일 등)이 있는지 확인'],
    ['· 서명이 안 붙음 → 메뉴 [내 Gmail 서명 확인] 실행. 원인이 표시됩니다.'],
    ['· 같은 사람에게 다시 보내야 함 → 메뉴 [발송기록 초기화] (그 달 중복방지가 풀립니다)'],
    ['· 발송 한도 → 메뉴 [남은 일일 발송량 확인]. 워크스페이스 계정은 하루 약 1,500통.'],
    ['· 숨겨진 "_발송기록" 시트는 중복발송 방지용입니다. 지우지 마세요.'],
    [''],
    ['최초 작성일: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')]
  ];

  sh.getRange(1, 1, lines.length, 1).setValues(lines);
  sh.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  for (var i = 0; i < lines.length; i++) {
    if (String(lines[i][0]).indexOf(H) === 0) {
      sh.getRange(i + 1, 1).setFontWeight('bold').setFontSize(11).setBackground('#eef3fb');
    }
  }
  sh.setColumnWidth(1, 900);
  sh.getRange(1, 1, lines.length, 1).setWrap(true);
  return sh;
}
