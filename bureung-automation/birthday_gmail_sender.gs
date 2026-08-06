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
var ERP_SHEET = 'ERP최신본';   // ERP 사원명부 export를 붙여넣는 시트
var DIFF_SHEET = '대조결과';   // 명단 대조 변경 내역 대시보드

// 명단 대조에서 변경을 감지할 항목 (명단 컬럼 키 ↔ ERP 머리글 후보)
var TRACK = [
  { key: 'status', label: '재직여부', cands: ['재직여부','재직상태'] },
  { key: 'rank',   label: '직급',     cands: ['직급'] },
  { key: 'role',   label: '직책',     cands: ['직책'] },
  { key: 'dept',   label: '비용센터', cands: ['비용센터','부서'] },
  { key: 'email',  label: '이메일',   cands: ['이메일','메일'] }
];

// [설정] 시트가 없을 때 사용되는 기본값 (설정 시트가 있으면 그쪽이 우선)
var DEFAULTS = {
  senderName: '부릉 피플실',
  subject: '[피플실] {연도}년 {월}월 생일휴가 부여 및 사용 안내(~{말일짧게}까지)',
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
    .addItem('⓪ 명단 대조(ERP 최신본 반영)', 'syncRoster')
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
  var skip = [SETTINGS_SHEET, GUIDE_SHEET, LOG_SHEET, ERP_SHEET, DIFF_SHEET];
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
               '재직여부','생년월일','이메일','메일','생일월',
               '대상 여부','제외 사유','발송하기','발송 여부','상태'];
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
    reason:   findCol_(h, ['제외 사유','제외사유'], true),          // 없으면 사유를 '대상 여부'에 함께 표기
    sendFlag: findCol_(h, ['발송하기','발송 여부','발송여부'], true),
    result:   findCol_(h, ['상태','발송상태','발송 상태'], true)     // 발송 결과 기록
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
  var html = bodyToHtml_(body);
  var text = body;
  if (sig.ok) {
    // Gmail이 수동 작성 시 넣는 것과 동일한 서명 구분선(--)
    html += '<div style="margin-top:22px;">--</div>' +
            '<div style="margin-top:10px;">' + sig.signature + '</div>';
    text += '\n\n--\n';
  }
  GmailApp.sendEmail(email, subject, text, { htmlBody: html, name: cfg.senderName });
}


// ===== ⓪ 명단 대조 (ERP 최신본 반영) =====
// 매월 발송 전, ERP 사원명부 최신본과 대조해
//  (1) 재직여부 · 직급 · 직책 · 비용센터 · 이메일 변경분을 명단에 반영
//  (2) 퇴직 추정자와 임원은 '비대상' 처리
//  (3) 신규 입사자를 명단 아래에 추가
//  (4) 변경 내역을 [대조결과] 시트에 기록하고, 명단 최종 수정일을 오늘로 갱신
function syncRoster() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var erpSheet = ss.getSheetByName(ERP_SHEET);
  if (!erpSheet) {
    ui.alert("'" + ERP_SHEET + "' 시트가 없습니다.\n\n" +
      '1) 시트 하단 + 를 눌러 새 시트를 만들고 이름을 "' + ERP_SHEET + '" 으로 바꾸세요.\n' +
      '2) ERP(옴니이솔) 사원명부 export를 머리글째 그대로 붙여넣으세요.\n' +
      '3) 다시 이 메뉴를 실행하세요.');
    return;
  }

  var cfg = readSettings_();
  var p = parse_();
  var col = p.col, values = p.values, sheet = p.sheet;
  var header = values[p.headerRow];

  var eVals = erpSheet.getDataRange().getValues();
  if (eVals.length < 2) { ui.alert("'" + ERP_SHEET + "' 시트에 데이터가 없습니다."); return; }
  var eHeaderRow = detectHeaderRow_(eVals);
  var eHeader = eVals[eHeaderRow];
  var eCol = {
    empno: findCol_(eHeader, ['사원번호','사번']),
    name:  findCol_(eHeader, ['성명','이름'])
  };
  TRACK.forEach(function (t) { eCol[t.key] = findCol_(eHeader, t.cands); });
  if (eCol.empno < 0) { ui.alert("'" + ERP_SHEET + "' 시트에서 '사원번호' 컬럼을 찾을 수 없습니다."); return; }

  // ERP 인덱스
  var erp = {}, erpOrder = [];
  for (var er = eHeaderRow + 1; er < eVals.length; er++) {
    var no = String(eVals[er][eCol.empno] || '').trim();
    if (!no) continue;
    erp[no] = eVals[er];
    erpOrder.push(no);
  }

  var isExec = function (v) { return v && cfg.excludeRanks.indexOf(v) >= 0; };

  // 명단 대조
  var inRoster = {}, changes = [], gone = [], noEmpno = [];
  for (var r = p.headerRow + 1; r < values.length; r++) {
    var row = values[r];
    var nm = col.name >= 0 ? String(row[col.name] || '').trim() : '';
    var eno = col.empno >= 0 ? String(row[col.empno] || '').trim() : '';
    if (!nm && !eno) continue;
    if (!eno) { noEmpno.push({ rowNum: r + 1, name: nm }); continue; }
    inRoster[eno] = true;
    if (!erp[eno]) {
      var curSt = col.status >= 0 ? String(row[col.status] || '').trim() : '';
      if (curSt !== '퇴사') gone.push({ rowNum: r + 1, empno: eno, name: nm, from: curSt });
      continue;
    }
    TRACK.forEach(function (t) {
      var rc = col[t.key], ec = eCol[t.key];
      if (rc < 0 || ec < 0) return;
      var cur = String(row[rc] == null ? '' : row[rc]).trim();
      var neu = String(erp[eno][ec] == null ? '' : erp[eno][ec]).trim();
      if (neu && cur !== neu) {
        changes.push({ rowNum: r + 1, empno: eno, name: nm, key: t.key, label: t.label,
                       from: cur, to: neu, col: rc });
      }
    });
  }

  // 신규 입사자
  var added = [];
  for (var k = 0; k < erpOrder.length; k++) if (!inRoster[erpOrder[k]]) added.push(erpOrder[k]);

  if (!changes.length && !gone.length && !added.length) {
    stampUpdated_(sheet);
    ui.alert('명단 대조 완료\n\n변경 사항이 없습니다. (ERP 최신본과 일치)' +
      (noEmpno.length ? '\n\n※ 사원번호가 없어 대조에서 제외한 행: ' + noEmpno.length + '개' : '') +
      '\n\n명단 최종 수정일을 오늘로 갱신했습니다.');
    return;
  }

  // 확인 메시지
  var byLabel = {};
  changes.forEach(function (c) { byLabel[c.label] = (byLabel[c.label] || 0) + 1; });
  var brief = function (arr, fmt) {
    return arr.slice(0, 6).map(fmt).join(', ') + (arr.length > 6 ? ' 외 ' + (arr.length - 6) + '명' : '');
  };
  var msg = '명단 대조 결과\n\n';
  if (changes.length) {
    msg += '· 정보 변경 ' + changes.length + '건\n';
    Object.keys(byLabel).forEach(function (L) { msg += '   - ' + L + ' ' + byLabel[L] + '건\n'; });
    msg += '   ' + brief(changes, function (c) { return c.name + '(' + c.label + ')'; }) + '\n';
  }
  if (gone.length) msg += '· ERP에 없음(퇴사 추정) ' + gone.length + '명\n   ' +
    brief(gone, function (x) { return x.name; }) + '\n';
  if (added.length) msg += '· 신규 입사자 ' + added.length + '명\n   ' +
    brief(added, function (no) { return String(erp[no][eCol.name] || no); }) + '\n';
  if (noEmpno.length) msg += '· 사원번호 없음(대조 제외) ' + noEmpno.length + '개\n';
  msg += '\n반영할까요?\n' +
         '- 변경된 값을 ERP 기준으로 갱신\n' +
         '- 퇴사 추정자 · 휴직자 · 임원은 "비대상" 처리\n' +
         '- 신규 입사자를 명단 맨 아래에 추가\n' +
         '- 변경 내역을 [' + DIFF_SHEET + '] 시트에 기록';

  if (ui.alert('명단 대조', msg, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  var now = new Date();
  var stamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var log = [];

  // (1) 변경분 반영
  var touched = {};
  changes.forEach(function (c) {
    sheet.getRange(c.rowNum, c.col + 1).setValue(c.to);
    touched[c.rowNum] = true;
    log.push([stamp, '정보변경', c.empno, c.name, c.label, c.from || '(없음)', c.to]);
  });

  // 변경된 행의 대상 여부 재평가 (재직 아님 · 임원이면 비대상)
  Object.keys(touched).forEach(function (rn) {
    var rowNum = parseInt(rn, 10);
    var st = col.status >= 0 ? String(sheet.getRange(rowNum, col.status + 1).getValue() || '').trim() : '';
    var rk = col.rank >= 0 ? String(sheet.getRange(rowNum, col.rank + 1).getValue() || '').trim() : '';
    var ro = col.role >= 0 ? String(sheet.getRange(rowNum, col.role + 1).getValue() || '').trim() : '';
    var bad = (st && cfg.statuses.indexOf(st) < 0) || isExec(rk) || isExec(ro);
    if (bad) {
      if (col.targetYN >= 0) sheet.getRange(rowNum, col.targetYN + 1).setValue('비대상');
      if (col.sendFlag >= 0) sheet.getRange(rowNum, col.sendFlag + 1).setValue(false);
    }
  });

  // (2) 퇴사 추정자
  gone.forEach(function (x) {
    if (col.status >= 0) sheet.getRange(x.rowNum, col.status + 1).setValue('퇴사');
    if (col.targetYN >= 0) sheet.getRange(x.rowNum, col.targetYN + 1).setValue('비대상');
    if (col.sendFlag >= 0) sheet.getRange(x.rowNum, col.sendFlag + 1).setValue(false);
    log.push([stamp, '퇴사추정', x.empno, x.name, '재직여부', x.from || '(없음)', '퇴사']);
  });

  // (3) 신규 입사자 추가 — 명단 머리글과 같은 이름의 ERP 열을 옮겨 담습니다.
  if (added.length) {
    var lastNo = 0;
    var noCol = findCol_(header, ['No','no','번호'], true);
    if (noCol >= 0) {
      for (var q = values.length - 1; q > p.headerRow; q--) {
        var v = parseInt(values[q][noCol], 10);
        if (!isNaN(v)) { lastNo = v; break; }
      }
    }
    var newRows = added.map(function (no, i) {
      var src = erp[no];
      var out = [];
      for (var c = 0; c < header.length; c++) {
        var hName = String(header[c] || '').trim();
        if (noCol >= 0 && c === noCol) { out.push(lastNo + i + 1); continue; }
        if (c === col.bmonth) {
          var bi = findCol_(eHeader, ['생년월일','생일','출생']);
          out.push(bi >= 0 ? (monthFromBirth_(src[bi]) || '') : '');
          continue;
        }
        if (c === col.targetYN || c === col.reason || c === col.sendFlag || c === col.result) { out.push(''); continue; }
        var ei = hName ? findCol_(eHeader, [hName], true) : -1;
        out.push(ei >= 0 ? src[ei] : '');
      }
      log.push([stamp, '신규입사', no, String(src[eCol.name] || ''), '-', '-', '명단 추가']);
      return out;
    });
    sheet.getRange(values.length + 1, 1, newRows.length, header.length).setValues(newRows);
  }

  // (4) 대시보드 기록 + 최종 수정일 갱신
  writeDiffSheet_(ss, log, stamp, { changed: changes.length, gone: gone.length, added: added.length });
  stampUpdated_(sheet);

  ui.alert('반영 완료\n\n' +
    '· 정보 변경 ' + changes.length + '건\n' +
    '· 퇴사 표시 ' + gone.length + '명\n' +
    '· 신규 추가 ' + added.length + '명\n\n' +
    "변경 내역은 ['" + DIFF_SHEET + "'] 시트에서 확인하세요.\n" +
    '이어서 [① 대상 분류]를 실행하세요.');
}

// 명단 최상단의 '명단 최종 수정일' 값을 오늘 날짜로 갱신 (라벨 오른쪽 두 번째 칸 = C열)
function stampUpdated_(sheet) {
  var rng = sheet.getRange(1, 1, Math.min(3, sheet.getLastRow()), 4).getValues();
  for (var r = 0; r < rng.length; r++) {
    for (var c = 0; c < rng[r].length; c++) {
      if (String(rng[r][c] || '').indexOf('명단 최종 수정일') >= 0) {
        var cell = sheet.getRange(r + 1, c + 3);   // 라벨이 A열이면 C열
        cell.setValue(new Date());
        cell.setNumberFormat('yyyy-mm-dd');
        return;
      }
    }
  }
}

// 변경 내역 대시보드 — 최신 대조 결과가 항상 맨 위에 쌓입니다.
function writeDiffSheet_(ss, log, stamp, sum) {
  var sh = ss.getSheetByName(DIFF_SHEET);
  var HEAD = ['대조일시', '구분', '사원번호', '성명', '변경 항목', '이전', '변경 후'];
  if (!sh) {
    sh = ss.insertSheet(DIFF_SHEET);
    sh.getRange(1, 1).setValue('명단 대조 결과 — 최신 대조가 맨 위에 표시됩니다');
    sh.getRange(1, 1, 1, HEAD.length).merge().setFontWeight('bold').setBackground('#fff3cd');
    sh.getRange(2, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#eeeeee');
    sh.setFrozenRows(2);
    [110, 80, 105, 80, 85, 175, 175].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  }
  var title = stamp + ' 대조  —  정보변경 ' + sum.changed + '건 · 퇴사추정 ' + sum.gone +
              '명 · 신규입사 ' + sum.added + '명';
  var block = [[title, '', '', '', '', '', '']].concat(log);
  sh.insertRowsBefore(3, block.length);
  sh.getRange(3, 1, block.length, HEAD.length).setValues(block);
  sh.getRange(3, 1, 1, HEAD.length).merge().setFontWeight('bold').setBackground('#e7f4e4');
  sh.getRange(4, 1, Math.max(log.length, 1), HEAD.length).setBackground(null);
  sh.activate();
}


// ===== ① 대상 분류 =====
function markTargets() {
  var ui = SpreadsheetApp.getUi();
  var cfg = readSettings_();
  var p = parse_();
  var col = p.col, values = p.values, sheet = p.sheet;
  if (col.targetYN < 0 || col.sendFlag < 0) {
    ui.alert("명단 시트에 '대상 여부' 와 '발송하기' 컬럼이 필요합니다.\n" +
             "(인식된 머리글 행: " + (p.headerRow + 1) + '행)'); return;
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
    if (!reason && !isValidEmail_(email)) reason = '이메일 없음';
    if (reason) {
      // '제외 사유' 열이 없으면 사유를 '대상 여부'에 함께 적습니다.
      yn.push([col.reason >= 0 ? '비대상' : '비대상(' + reason + ')']);
      rs.push([reason]); sf.push([false]);
      if (reason === '이메일 없음') cN++; else cE++;
    } else {
      yn.push(['대상']); rs.push(['']); sf.push([true]); cT++;
    }
  }

  sheet.getRange(first + 1, col.targetYN + 1, n, 1).setValues(yn);
  if (col.reason >= 0) sheet.getRange(first + 1, col.reason + 1, n, 1).setValues(rs);
  var sfRange = sheet.getRange(first + 1, col.sendFlag + 1, n, 1);
  sfRange.insertCheckboxes();
  sfRange.setValues(sf);

  ui.alert(month + '월 분류 완료\n- 대상 ' + cT + '명(발송하기 체크됨) · 제외 ' + cE + '명 · 이메일없음 ' + cN + '명\n\n' +
    "'발송하기' 열을 확인/수정한 뒤, 메뉴 ③으로 발송하세요.");
}


// ===== ③ 발송 =====
function sendMarked() {
  var ui = SpreadsheetApp.getUi();
  var cfg = readSettings_();
  var p = parse_();
  var col = p.col, values = p.values, sheet = p.sheet;
  if (col.sendFlag < 0) { ui.alert("'발송하기' 컬럼을 찾을 수 없습니다. 먼저 ①을 실행하세요."); return; }

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
    why += "· '발송하기'가 체크된 행: " + stat.checked + '명\n';
    if (stat.noEmail) why += '· 이메일이 없거나 형식이 잘못되어 제외: ' + stat.noEmail + '명\n';
    if (stat.already) {
      why += '· 이번 달에 이미 발송되어 제외(중복 방지): ' + stat.already + '명\n';
      why += '   → ' + alreadyNames.slice(0, 10).join(', ') + (alreadyNames.length > 10 ? ' 외' : '') + '\n';
      why += "\n같은 사람에게 다시 보내려면 메뉴 [발송기록 초기화]를 실행한 뒤 다시 시도하세요.";
    }
    if (!stat.checked) why += "\n먼저 [① 대상 분류]를 실행하거나, '발송하기' 열을 직접 체크하세요.";
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
      // 결과는 '상태' 열에 기록(없으면 '대상 여부' 열에)
      var outCol = col.result >= 0 ? col.result : col.targetYN;
      if (outCol >= 0) sheet.getRange(t.rowNum, outCol + 1).setValue('발송완료');
      sheet.getRange(t.rowNum, col.sendFlag + 1).setValue(false);
      ok++;
    } catch (e) {
      fail++; fails.push(t.name + ': ' + e.message);
      if (col.result >= 0) sheet.getRange(t.rowNum, col.result + 1).setValue('발송실패');
    }
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
    ['생일휴가 안내 업무 – 사용법 및 인수인계 문서'],
    [''],
    [H + ' 이 도구가 하는 일 / 하지 않는 일  (먼저 읽어주세요)'],
    ['[하는 일]  대상자 선별 · 명단 최신화 · 안내 메일 작성과 발송 · 발송 이력 관리'],
    ['[하지 않는 일]  ERP(옴니이솔)에 생일휴가를 실제로 부여하거나 삭제하는 작업'],
    ['   → ERP 휴가 부여·삭제는 지금도 담당자가 직접 해야 합니다. 아래 연간 흐름을 참고하세요.'],
    [''],
    [H + ' 연간 흐름'],
    ['[연 1회 · 연초]  전 재직자에게 생일휴가 1일을 ERP에서 일괄 부여합니다.'],
    ['   옴니이솔 > 근태일수등록 > 사원별일수등록 에서 근태코드 "생일휴가" 선택'],
    ['   시작일 = 생일월 1일 / 종료일 = 생일월 말일 / 부여일수 = 1'],
    ['   ※ 이때 임원은 부여 대상에서 제외해야 합니다. (2026년에는 누락되어 이후 개별 삭제함)'],
    ['[매월 말]  다음 달 생일자를 확정하고 안내를 발송합니다. 아래 순서를 따르세요.'],
    [''],
    [H + ' 매월 작업 순서'],
    ['1. ERP에서 사원명부(재직자)를 export 합니다.'],
    ['2. [ERP최신본] 시트에 머리글째 그대로 붙여넣습니다.'],
    ['   ※ 열 순서를 맞출 필요 없습니다. 항목 이름으로 자동 인식합니다.'],
    ['   ※ 반드시 전체 명부를 받으세요. 일부만 받으면 재직자가 퇴사로 표시됩니다.'],
    ['3. 메뉴 [⓪ 명단 대조(ERP 최신본 반영)] 실행'],
    ['   → 재직여부 · 직급 · 직책 · 비용센터 · 이메일 변경분이 명단에 반영됩니다.'],
    ['   → 휴직 · 퇴직자와 임원은 자동으로 "비대상" 처리됩니다.'],
    ['   → 신규 입사자는 명단 맨 아래에 추가됩니다(생일월 자동 계산).'],
    ['   → 무엇이 바뀌었는지는 [대조결과] 시트에서 확인할 수 있습니다.'],
    ['   → 명단 최상단의 "명단 최종 수정일"이 실행일로 갱신됩니다.'],
    ['4. ERP 휴가 부여를 조정합니다. (이 부분은 수동)'],
    ['   - 신규 입사자 중 다음 달 생일자 → ERP에서 생일휴가 1일 부여'],
    ['   - 퇴직 · 휴직으로 바뀐 사람 → ERP에서 부여 내역 삭제'],
    ['   - 임원으로 승진한 사람 → ERP에서 부여 내역 삭제'],
    ['   ※ [대조결과] 시트를 보면 누구를 조정해야 하는지 바로 알 수 있습니다.'],
    ['5. 메뉴 [① 대상 분류(월 선택)] → 안내할 달 입력'],
    ['   → "대상 여부"와 "발송하기" 열이 자동으로 채워집니다.'],
    ['6. 대상자 목록을 상급자에게 전달해 확인받습니다.'],
    ['7. 메뉴 [② 나에게 테스트 발송] → 문구와 서명을 확인합니다.'],
    ['8. 메뉴 [③ 체크된 사람에게 발송] → 발송된 행의 "상태"가 "발송완료"로 바뀝니다.'],
    [''],
    [H + ' 문구 · 규칙을 바꾸고 싶을 때'],
    ['[설정] 시트에서 수정하세요. 코드(Apps Script)는 열 필요 없습니다.'],
    ['  - 메일 제목 / 본문 : {이름} {월} {말일} {연도} {말일짧게} 가 자동으로 채워집니다.'],
    ['  - 제외 직급/직책 : 여기 적힌 직급·직책은 발송 대상에서 자동 제외됩니다(임원 판별 기준).'],
    ['  - 제외 부서 / 발송 대상 재직상태 / 이름 표기'],
    ['  ※ 임원 판별 기준은 인사 규정과 일치하는지 주기적으로 확인이 필요합니다.'],
    [''],
    [H + ' ★ 담당자가 바뀔 때 반드시 할 일'],
    ['이 도구는 "실행하는 사람의 구글 계정"으로 메일을 보냅니다.'],
    ['담당자가 바뀌면 아래를 진행하세요. 안 하면 어느 날 갑자기 멈춥니다.'],
    ['1. 스프레드시트 소유권을 후임자(또는 팀 공용 계정)에게 이전'],
    ['   [공유] → 후임자 추가 → 점 3개 → "소유권 이전"'],
    ['2. 후임자가 [확장 프로그램] → [Apps Script] 진입'],
    ['3. 편집기 왼쪽 [서비스] 옆 + → Gmail 선택 → 추가 (서명 사용 시 필수)'],
    ['4. [② 나에게 테스트 발송] 실행 → 권한 승인("고급" → "이동" → "허용")'],
    ['5. 테스트 메일이 정상 도착하면 완료'],
    ['※ 메일 서명은 "실행하는 사람의 Gmail 서명"이 자동으로 붙습니다.'],
    ['※ 연초 일괄 부여(위 연간 흐름)는 1년에 한 번뿐이라 놓치기 쉽습니다. 반드시 인수인계하세요.'],
    [''],
    [H + ' 문제 해결'],
    ['· "이메일 컬럼을 찾을 수 없습니다" → 명단 머리글(성명/이메일 등)이 있는지 확인'],
    ['· "발송할 대상이 없습니다" → 안내창에 이유가 표시됩니다(체크 안 됨 / 이미 발송 등)'],
    ['· 서명이 안 붙음 → 메뉴 [내 Gmail 서명 확인] 실행 (원인이 표시됩니다)'],
    ['· 같은 사람에게 다시 보내야 함 → 메뉴 [발송기록 초기화]'],
    ['· 숨겨진 "_발송기록" 시트는 중복발송 방지용입니다. 지우지 마세요.'],
    ['· [대조결과] 시트는 과거 대조 이력이 계속 쌓입니다. 지우지 않아도 됩니다.'],
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
