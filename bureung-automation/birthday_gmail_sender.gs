/**
 * 생일휴가 안내 – 구글 스프레드시트 → Gmail 자동 발송
 * ---------------------------------------------------------------
 * 별도 API/토큰/외부 서버 없이, 구글 스프레드시트에 내장된
 * Apps Script(GmailApp)로 안내 메일을 발송합니다.
 * 최초 1회 본인 구글 계정 권한 승인만 하면 됩니다.
 *
 * [설치]
 *  1) 명단 스프레드시트 열기 → 확장 프로그램 → Apps Script
 *  2) 이 코드 전체를 붙여넣고 저장
 *  3) 스프레드시트를 새로고침하면 상단에 "🎂 생일휴가 안내" 메뉴가 생깁니다
 *
 * [시트 형식] 1행은 머리글, 아래는 아래 컬럼을 인식합니다(이름/순서 유연):
 *  - 이름   : 이름, 성명, name
 *  - 이메일 : 이메일, 메일, email, gmail  (필수)
 *  - 생일월 : 생일월, 월  (1~12 숫자. 비어 있으면 '해당 월 필터 없이 전체' 대상)
 *  - 발송상태 / 발송일시 : 없으면 자동으로 만들어 기록합니다(중복 발송 방지)
 */

// ===== 설정 =====
const CONFIG = {
  SHEET_NAME: '',                 // 비우면 활성 시트 사용. 특정 시트명 지정 가능(예: '명단')
  SENDER_NAME: '부릉 피플실',      // 받는 사람에게 보이는 발신자 이름
  SUBJECT: '[부릉] {월}월 생일 축하 & 생일휴가 안내 🎂',
  RESEND_ALREADY_SENT: false      // true면 '발송완료' 행도 다시 보냅니다(재발송)
};

// 메시지 본문 템플릿 ({이름} {월} {말일} 자동 치환)
const BODY_TEXT =
`안녕하세요, {이름}님.
{월}월 생일을 축하드립니다!🎉
부릉에서는 생일 당월, 더 행복한 한 달 보내실 수 있도록 생일휴가 1일을 드리고 있습니다.

📅 생일휴가 1일 부여
- 신청 및 사용 기한: 생일 당월인 {월}월 말일({말일})까지
- 말일 이후 자동소멸되며, 당월 내에서만 사용 가능

📌 신청 방법
- 옴니이솔 → 인사관리 → My HR(ESS) → 근태신청(NEW) → [휴가신청서]
  → 휴가종류 선택: 기타휴가 > 생일휴가 선택

📢 결재 라인
- 1차 조직장(결재) - 피플실(합의)
- 주의사항: 반드시 '피플실(합의)' 지정

즐겁고 행복한 한 달 되세요! 🎉`;


// ===== 메뉴 =====
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎂 생일휴가 안내')
    .addItem('① 미리보기 (첫 대상자)', 'previewFirst')
    .addItem('② 나에게 테스트 발송', 'sendTestToMe')
    .addSeparator()
    .addItem('③ 이번 달 생일자에게 발송', 'sendBirthdayEmails')
    .addSeparator()
    .addItem('발송 상태 초기화', 'resetStatus')
    .addItem('남은 일일 발송량 확인', 'showQuota')
    .addToUi();
}


// ===== 유틸 =====
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return CONFIG.SHEET_NAME ? ss.getSheetByName(CONFIG.SHEET_NAME) : ss.getActiveSheet();
}

function lastDayOfMonth_(year, month) { return new Date(year, month, 0).getDate(); }

function findCol_(headers, candidates) {
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i]).trim().toLowerCase();
    if (candidates.some(c => h.includes(c.toLowerCase()))) return i;
  }
  return -1;
}

// 시트를 읽어 컬럼 인덱스와 데이터를 정리. 상태/일시 컬럼이 없으면 추가.
function readSheet_() {
  const sheet = getSheet_();
  if (!sheet) throw new Error('시트를 찾을 수 없습니다. CONFIG.SHEET_NAME을 확인하세요.');
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length < 2) throw new Error('데이터가 없습니다. 1행 머리글 아래에 명단을 입력하세요.');

  let headers = values[0].map(h => String(h).trim());
  const col = {
    name:  findCol_(headers, ['이름', '성명', 'name']),
    email: findCol_(headers, ['이메일', '메일', 'email', 'gmail']),
    month: findCol_(headers, ['생일월', '월', 'month']),
    status: findCol_(headers, ['발송상태', '상태', 'status']),
    sentAt: findCol_(headers, ['발송일시', '일시', 'sent'])
  };
  if (col.email < 0) throw new Error("'이메일' 컬럼을 찾을 수 없습니다. 머리글에 이메일 컬럼을 추가하세요.");

  // 상태/일시 컬럼이 없으면 맨 뒤에 생성
  let addCol = headers.length;
  if (col.status < 0) { sheet.getRange(1, addCol + 1).setValue('발송상태'); col.status = addCol; addCol++; }
  if (col.sentAt < 0) { sheet.getRange(1, addCol + 1).setValue('발송일시'); col.sentAt = addCol; addCol++; }

  return { sheet, values, col };
}

function buildMessage_(name, month) {
  const year = new Date().getFullYear();
  const last = lastDayOfMonth_(year, month);
  const lastLabel = `${month}월 ${last}일`;
  const rep = s => String(s)
    .replace(/{이름}/g, name)
    .replace(/{월}/g, month)
    .replace(/{말일}/g, lastLabel);
  const text = rep(BODY_TEXT);
  const subject = rep(CONFIG.SUBJECT);
  return { subject, text, html: textToHtml_(text) };
}

// 간단한 텍스트 → HTML 변환(줄바꿈/섹션 헤더 강조)
function textToHtml_(text) {
  const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const body = text.split('\n').map(line => {
    const t = line.trim();
    if (/^[📅📌📢]/.test(t)) return `<div style="margin:16px 0 6px;font-weight:600;">${esc(line)}</div>`;
    if (t === '') return '<div style="height:8px;"></div>';
    if (t.startsWith('- ')) return `<div style="padding-left:14px;">• ${esc(t.slice(2))}</div>`;
    return `<div>${esc(line)}</div>`;
  }).join('');
  return `<div style="font-family:'Apple SD Gothic Neo','Malgun Gothic',sans-serif;font-size:15px;line-height:1.7;color:#1a1a18;max-width:560px;">${body}</div>`;
}


// ===== 대상자 선별 =====
function collectTargets_(targetMonth) {
  const { sheet, values, col } = readSheet_();
  const targets = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const email = String(row[col.email] || '').trim();
    if (!email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) continue;
    const name = col.name >= 0 ? String(row[col.name] || '').trim() : '';
    const month = col.month >= 0 ? parseInt(row[col.month], 10) : targetMonth;
    // 생일월 컬럼이 있으면 이번 달만, 없으면 전체
    if (col.month >= 0 && month !== targetMonth) continue;
    const status = String(row[col.status] || '').trim();
    const already = status === '발송완료';
    targets.push({ rowIndex: r + 1, name, email, month: month || targetMonth, already });
  }
  return { sheet, col, targets };
}


// ===== 미리보기 / 테스트 =====
function previewFirst() {
  const ui = SpreadsheetApp.getUi();
  const month = new Date().getMonth() + 1;
  const { targets } = collectTargets_(month);
  if (!targets.length) { ui.alert(`${month}월 생일자(유효 이메일)가 없습니다.`); return; }
  const t = targets[0];
  const m = buildMessage_(t.name || '홍길동', t.month);
  ui.alert(`미리보기 — ${t.name} <${t.email}>\n제목: ${m.subject}\n\n${m.text}`);
}

function sendTestToMe() {
  const ui = SpreadsheetApp.getUi();
  const me = Session.getActiveUser().getEmail();
  if (!me) { ui.alert('본인 이메일을 확인할 수 없습니다.'); return; }
  const month = new Date().getMonth() + 1;
  const m = buildMessage_('테스트', month);
  GmailApp.sendEmail(me, '[테스트] ' + m.subject, m.text, { htmlBody: m.html, name: CONFIG.SENDER_NAME });
  ui.alert(`테스트 메일을 ${me} 로 보냈습니다. 받은편지함을 확인하세요.`);
}


// ===== 실제 발송 =====
function sendBirthdayEmails() {
  const ui = SpreadsheetApp.getUi();
  const nowMonth = new Date().getMonth() + 1;

  // 발송 대상 월 입력(기본: 이번 달)
  const resp = ui.prompt('생일휴가 안내 발송',
    `발송할 대상 월을 입력하세요 (1~12).\n기본값: ${nowMonth}월`, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const input = resp.getResponseText().trim();
  const month = input ? parseInt(input, 10) : nowMonth;
  if (!(month >= 1 && month <= 12)) { ui.alert('1~12 사이의 월을 입력하세요.'); return; }

  const { sheet, col, targets } = collectTargets_(month);
  const toSend = targets.filter(t => CONFIG.RESEND_ALREADY_SENT || !t.already);
  const skipped = targets.length - toSend.length;

  if (!toSend.length) {
    ui.alert(`${month}월 발송 대상이 없습니다.` + (skipped ? `\n(이미 발송완료 ${skipped}명 제외)` : ''));
    return;
  }

  const quota = MailApp.getRemainingDailyQuota();
  if (quota < toSend.length) {
    ui.alert(`일일 발송 한도가 부족합니다.\n대상 ${toSend.length}명 / 남은 한도 ${quota}통.\n한도는 다음 날 초기화됩니다.`);
    return;
  }

  const confirm = ui.alert('발송 확인',
    `${month}월 생일자 ${toSend.length}명에게 안내 메일을 보냅니다.` +
    (skipped ? `\n(이미 발송완료 ${skipped}명 제외)` : '') +
    `\n\n계속할까요?`, ui.ButtonSet.OK_CANCEL);
  if (confirm !== ui.Button.OK) return;

  let ok = 0, fail = 0;
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  toSend.forEach(t => {
    try {
      const m = buildMessage_(t.name || '', t.month);
      GmailApp.sendEmail(t.email, m.subject, m.text, { htmlBody: m.html, name: CONFIG.SENDER_NAME });
      sheet.getRange(t.rowIndex, col.status + 1).setValue('발송완료');
      sheet.getRange(t.rowIndex, col.sentAt + 1).setValue(stamp);
      ok++;
    } catch (e) {
      sheet.getRange(t.rowIndex, col.status + 1).setValue('실패: ' + e.message);
      fail++;
    }
    Utilities.sleep(300); // 과도한 연속 호출 방지
  });

  ui.alert(`발송 완료\n성공 ${ok}명 · 실패 ${fail}명\n남은 일일 한도 약 ${MailApp.getRemainingDailyQuota()}통`);
}


// ===== 관리 =====
function resetStatus() {
  const ui = SpreadsheetApp.getUi();
  const c = ui.alert('발송 상태 초기화', '모든 행의 발송상태/발송일시를 지웁니다. 계속할까요?', ui.ButtonSet.OK_CANCEL);
  if (c !== ui.Button.OK) return;
  const { sheet, values, col } = readSheet_();
  for (let r = 1; r < values.length; r++) {
    sheet.getRange(r + 1, col.status + 1).clearContent();
    sheet.getRange(r + 1, col.sentAt + 1).clearContent();
  }
  ui.alert('초기화했습니다.');
}

function showQuota() {
  SpreadsheetApp.getUi().alert(`오늘 남은 Gmail 발송 한도: 약 ${MailApp.getRemainingDailyQuota()}통\n(개인 계정 500 / 워크스페이스 1,500 기준, 다음 날 초기화)`);
}
