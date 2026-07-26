/**
 * ============================================================================
 *  부릉 사업소득 수정신고 — ERP(옴니이솔/icube) 업로드파일 자동 생성기
 * ----------------------------------------------------------------------------
 *  '수정신고_부릉' 시트의 소득해당자 최종금액을 읽어서
 *  ERP 「사업기타이자배당소득엑셀업로드」 양식(TSMINC00700_F)으로 바꿔줍니다.
 *
 *  ▶ 설치 방법
 *    1) 대상 구글 스프레드시트 열기
 *    2) 상단 메뉴 [확장 프로그램] → [Apps Script]
 *    3) 이 파일 내용을 통째로 붙여넣고 저장(💾)
 *    4) 시트로 돌아와 새로고침 → 상단에 [부릉 자동화] 메뉴가 생김
 *
 *  ▶ 사용 방법
 *    - (선택) 처리할 행들을 드래그로 선택 → 선택한 행만 생성됨
 *      (아무것도 선택 안 하면 전체 데이터 행을 대상으로 함)
 *    - [부릉 자동화] → [① ERP 업로드파일 생성(소득해당자)]
 *    - 새 스프레드시트가 만들어지고 링크가 뜸
 *    - 그 파일을 열어 [파일] → [다운로드] → [Microsoft Excel(.xlsx)]
 *    - 옴니이솔 [엑셀파일업로드]에서 그대로 업로드 → 오류검증 → 서식반영
 *
 *  ⚠️ 확실치 않은 값은 전부 아래 CONFIG로 빼놨습니다. 실제와 다르면 여기만 고치세요.
 * ============================================================================
 */

var CONFIG = {
  // ── 원본 시트 ──────────────────────────────────────────────
  SOURCE_SHEET: '수정신고_부릉',
  DATA_START_ROW: 7,            // 데이터가 시작되는 행 (헤더 6행 다음)

  // ── 원본 시트의 열 위치 (1 = A열, 2 = B열 …). 구조 바뀌면 여기만 수정 ──
  COL: {
    부인자명:        9,   // I  소득부인자명 (값이 있으면 '새 그룹 시작'으로 인식)
    부인_주민번호:   10,  // J
    지급연월:        12,  // L
    부인_소득금액:   13,  // M
    부인_소득세:     16,  // P
    부인_지방세:     17,  // Q
    해당자명:        22,  // V  소득해당자(정정 대상자)명
    해당자_주민번호: 23,  // W
    해당자_기존_소득금액: 24, // X  (최종값이 비었을 때 폴백 계산용)
    해당자_기존_소득세:   25, // Y
    해당자_기존_지방세:   26, // Z
    최종_소득금액:   27,  // AA
    최종_소득세:     28,  // AB
    최종_지방세:     29,  // AC
    해당자_소득자코드: 0  // 시트에 소득자코드 열이 있으면 그 열번호를, 없으면 0
  },

  // ── 소득자코드(INCMPER_NO) 매핑 ─────────────────────────────
  //  ERP 양식엔 이름/주민번호 칸이 없어 '소득자코드'로만 소득자를 식별합니다.
  //  아래 이름의 시트를 만들어 [A열=주민번호, B열=소득자코드]로 채워두면
  //  주민번호를 보고 소득자코드를 자동으로 채워줍니다. (첫 행은 제목행)
  CODE_MAP_SHEET: '소득자코드매핑',

  // ── ERP 고정값 (매뉴얼 기준) ────────────────────────────────
  COMPANY_CD:  '1000',
  COMPANY_NM:  '(주)부릉',
  BIZR_NO:     '2068673707',
  BIZAREA_CD:  '1000',        // 사업장코드
  DEPT_CD:     'AE0000000',   // 부서코드
  BIZTP_FG_CD: '940918',      // 업종구분코드(퀵서비스)
  TAX_RT:      '3',           // 세율(%)

  // ── 지급액 기준 ─────────────────────────────────────────────
  //  'FINAL'  = 소득해당자 최종금액(AA/AB/AC) 을 업로드 → ERP 월소득을 덮어씀
  //             (매뉴얼 방법2 엑셀 예시가 최종액 42,100 을 보여줌 → 기본값)
  //  'DENIED' = 부인자에게서 이관되는 금액(M/P/Q) 만 업로드 → 기존에 '추가'
  //             (매뉴얼 방법1 방식)
  AMOUNT_BASIS: 'FINAL',

  OUTPUT_SHEET_NAME: 'TSMINC00700_F'  // ERP 업로드 양식 시트명 (건드리지 마세요)
};

// ERP 양식 컬럼 순서/제목 (실제 다운로드 양식과 100% 동일)
var ERP_CODES  = ['INCMPER_NO','PAY_YM','RVERS_YM','PAY_DT','BIZAREA_CD','DEPT_CD','WGS_AMT','BIZTP_FG_CD','TAX_RT','INTAX_AMT','LOCINTAX_AMT'];
var ERP_LABELS = ['소득자코드','지급연월','귀속년월','지급일','사업장코드','부서코드','지급액','업종구분코드','세율','소득세금액','지방소득세금액'];


/** 시트 열 때 커스텀 메뉴 생성 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('부릉 자동화')
    .addItem('① ERP 업로드파일 생성(소득해당자)', 'generateErpUpload')
    .addItem('② 부인자 삭제 대상 목록', 'generateDenialList')
    .addToUi();
}


/** ① 메인: ERP 업로드파일 생성 */
function generateErpUpload() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  if (!sh) { ui.alert('시트 없음', "'" + CONFIG.SOURCE_SHEET + "' 시트를 찾을 수 없습니다.", ui.ButtonSet.OK); return; }

  var lastRow = sh.getLastRow();
  var maxCol  = Math.max.apply(null, objValues(CONFIG.COL).concat([1]));
  var values  = sh.getRange(1, 1, lastRow, maxCol).getValues();

  // 선택 범위가 데이터 영역 안이면 그 행들만 대상으로
  var scope = getSelectionScope(sh, lastRow);

  var codeMap = loadCodeMap(ss);

  var out = [];
  var missingCode = [];   // 소득자코드 못 찾은 해당자
  var fallbackCnt = 0;    // 최종값이 비어 기존+부인으로 계산한 행 수
  var current = null;     // 현재 그룹의 소득해당자 정보

  for (var r = CONFIG.DATA_START_ROW; r <= lastRow; r++) {
    var row = values[r - 1];

    // 새 그룹 시작(부인자명이 있는 행)에서 소득해당자 정보 갱신
    if (notEmpty(cell(row, CONFIG.COL.부인자명))) {
      var rrn  = normalizeRrn(cell(row, CONFIG.COL.해당자_주민번호));
      var name = String(cell(row, CONFIG.COL.해당자명) || '').trim();
      var code = CONFIG.COL.해당자_소득자코드 > 0
                 ? String(cell(row, CONFIG.COL.해당자_소득자코드) || '').trim() : '';
      if (!code && rrn && codeMap[rrn]) code = codeMap[rrn];
      current = { rrn: rrn, name: name, code: code };
    }

    var ym = parseYm(cell(row, CONFIG.COL.지급연월));
    if (ym === null) continue;                                  // 금액 없는 행 skip
    if (scope && (r < scope.top || r > scope.bottom)) continue; // 선택 범위 밖 skip
    if (!current) continue;

    // 금액 산정 (DENIED = 부인금액 그대로 / FINAL = 최종값, 비면 기존+부인 폴백)
    var amt, intax, locint;
    if (CONFIG.AMOUNT_BASIS === 'DENIED') {
      amt    = toInt(cell(row, CONFIG.COL.부인_소득금액));
      intax  = toInt(cell(row, CONFIG.COL.부인_소득세)) || 0;
      locint = toInt(cell(row, CONFIG.COL.부인_지방세)) || 0;
    } else {
      amt    = toInt(cell(row, CONFIG.COL.최종_소득금액));
      intax  = toInt(cell(row, CONFIG.COL.최종_소득세));
      locint = toInt(cell(row, CONFIG.COL.최종_지방세));
      if (amt === null) {  // 최종값 공란 → 기존(해당자) + 부인 으로 계산
        amt    = (toInt(cell(row, CONFIG.COL.해당자_기존_소득금액)) || 0) + (toInt(cell(row, CONFIG.COL.부인_소득금액)) || 0);
        intax  = (toInt(cell(row, CONFIG.COL.해당자_기존_소득세))   || 0) + (toInt(cell(row, CONFIG.COL.부인_소득세))   || 0);
        locint = (toInt(cell(row, CONFIG.COL.해당자_기존_지방세))   || 0) + (toInt(cell(row, CONFIG.COL.부인_지방세))   || 0);
        if (amt > 0) fallbackCnt++;
      }
      intax  = intax  || 0;
      locint = locint || 0;
    }
    if (amt === null || amt <= 0) continue;  // 지급액 0/음수 행 skip

    if (!current.code) {
      missingCode.push((current.name || '(이름없음)') + ' / ' + (current.rrn || '(주민번호없음)'));
    }

    // ERP_CODES 순서대로: INCMPER_NO, PAY_YM, RVERS_YM, PAY_DT, BIZAREA_CD,
    //                     DEPT_CD, WGS_AMT, BIZTP_FG_CD, TAX_RT, INTAX_AMT, LOCINTAX_AMT
    out.push([
      current.code || '',        // INCMPER_NO (소득자코드)
      ym,                        // PAY_YM  (지급연월)  ⟵ 귀속=지급 동일 처리
      ym,                        // RVERS_YM(귀속년월)
      lastDayYmd(ym),            // PAY_DT  (지급일=해당월 말일)
      CONFIG.BIZAREA_CD,
      CONFIG.DEPT_CD,
      String(amt),               // WGS_AMT (지급액)
      CONFIG.BIZTP_FG_CD,
      CONFIG.TAX_RT,
      String(intax),             // INTAX_AMT
      String(locint)             // LOCINTAX_AMT
    ]);
  }

  if (out.length === 0) {
    ui.alert('생성할 데이터 없음', '조건에 맞는 소득해당자 행을 찾지 못했습니다.\n(지급연월·최종 소득금액이 채워진 행만 대상입니다.)', ui.ButtonSet.OK);
    return;
  }

  var url = writeErpFile(out);

  // 소득자코드 누락 경고
  var uniqMissing = uniq(missingCode);
  var msg = 'ERP 업로드 행 ' + out.length + '건 생성 완료 ✅\n\n' +
            '아래 파일을 열어 [파일]→[다운로드]→[Excel(.xlsx)] 후 옴니이솔에 업로드하세요.\n\n' + url;
  if (fallbackCnt > 0) {
    msg += '\n\nℹ️ ' + fallbackCnt + '건은 최종금액 칸이 비어 있어 (해당자 기존 + 부인금액)으로 자동 계산했습니다. 값을 한 번 확인해 주세요.';
  }
  if (uniqMissing.length > 0) {
    msg += '\n\n⚠️ 소득자코드(INCMPER_NO)가 비어 있는 대상 ' + uniqMissing.length + '명:\n · '
         + uniqMissing.slice(0, 20).join('\n · ')
         + (uniqMissing.length > 20 ? '\n … 외 ' + (uniqMissing.length - 20) + '명' : '')
         + "\n\n→ '" + CONFIG.CODE_MAP_SHEET + "' 시트에 [주민번호, 소득자코드]를 채우면 자동으로 채워집니다.";
  }
  ui.alert('완료', msg, ui.ButtonSet.OK);
}


/** ② 부인자 삭제 대상 목록 (수동 승인해제+삭제 시 참고용) */
function generateDenialList() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  if (!sh) { ui.alert('시트 없음', "'" + CONFIG.SOURCE_SHEET + "' 시트를 찾을 수 없습니다.", ui.ButtonSet.OK); return; }

  var lastRow = sh.getLastRow();
  var maxCol  = Math.max.apply(null, objValues(CONFIG.COL).concat([1]));
  var values  = sh.getRange(1, 1, lastRow, maxCol).getValues();
  var scope   = getSelectionScope(sh, lastRow);

  var out = [];
  var curName = '', curRrn = '';
  for (var r = CONFIG.DATA_START_ROW; r <= lastRow; r++) {
    var row = values[r - 1];
    if (notEmpty(cell(row, CONFIG.COL.부인자명))) {
      curName = String(cell(row, CONFIG.COL.부인자명) || '').trim();
      curRrn  = normalizeRrn(cell(row, CONFIG.COL.부인_주민번호));
    }
    var ym  = parseYm(cell(row, CONFIG.COL.지급연월));
    var amt = toInt(cell(row, CONFIG.COL.부인_소득금액));
    if (ym === null || amt === null || amt <= 0) continue;
    if (scope && (r < scope.top || r > scope.bottom)) continue;
    out.push([curName, curRrn, ym,
              String(amt),
              String(toInt(cell(row, CONFIG.COL.부인_소득세)) || 0),
              String(toInt(cell(row, CONFIG.COL.부인_지방세)) || 0)]);
  }
  if (out.length === 0) { ui.alert('대상 없음', '부인 대상 행을 찾지 못했습니다.', ui.ButtonSet.OK); return; }

  var newSs = SpreadsheetApp.create('부인자삭제목록_' + stamp());
  var ns = newSs.getActiveSheet();
  ns.setName('부인자삭제대상');
  var header = ['소득부인자명', '주민번호', '귀속/지급연월', '소득금액', '소득세', '지방소득세'];
  ns.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  ns.getRange(2, 1, out.length, header.length).setNumberFormat('@').setValues(out);
  ns.autoResizeColumns(1, header.length);
  ui.alert('완료', '부인자 삭제 대상 ' + out.length + '건 목록 생성 ✅\n\n' + newSs.getUrl(), ui.ButtonSet.OK);
}


/* ───────────────────────── 내부 유틸 ───────────────────────── */

/** ERP 양식 그대로 새 스프레드시트에 기록하고 URL 반환 */
function writeErpFile(dataRows) {
  var newSs = SpreadsheetApp.create('ERP업로드_' + stamp());
  var sh = newSs.getActiveSheet();
  sh.setName(CONFIG.OUTPUT_SHEET_NAME);

  // 회사 헤더 블록 (양식 1~3행) + 빈 4행 + 컬럼코드(5행) + 한글제목(6행)
  var head = [
    ['COMPANY_CD', 'COMPANY_NM', 'BIZR_NO'],
    ['회사코드', '회사명', '사업자번호'],
    [CONFIG.COMPANY_CD, CONFIG.COMPANY_NM, CONFIG.BIZR_NO]
  ];
  sh.getRange(1, 1, 3, 3).setNumberFormat('@').setValues(head);
  sh.getRange(5, 1, 1, ERP_CODES.length).setNumberFormat('@').setValues([ERP_CODES]);
  sh.getRange(6, 1, 1, ERP_LABELS.length).setNumberFormat('@').setValues([ERP_LABELS]);

  // 데이터 (7행부터) — ⚠️ 모든 셀 텍스트 서식
  sh.getRange(7, 1, dataRows.length, ERP_CODES.length)
    .setNumberFormat('@')
    .setValues(dataRows);

  sh.autoResizeColumns(1, ERP_CODES.length);
  SpreadsheetApp.flush();
  return newSs.getUrl();
}

/** '소득자코드매핑' 시트를 {주민번호: 소득자코드} 로 로드 */
function loadCodeMap(ss) {
  var map = {};
  var sh = ss.getSheetByName(CONFIG.CODE_MAP_SHEET);
  if (!sh) return map;
  var last = sh.getLastRow();
  if (last < 2) return map;
  var vals = sh.getRange(2, 1, last - 1, 2).getValues(); // A=주민번호, B=소득자코드
  for (var i = 0; i < vals.length; i++) {
    var rrn = normalizeRrn(vals[i][0]);
    var code = String(vals[i][1] || '').trim();
    if (rrn && code) map[rrn] = code;
  }
  return map;
}

/** 선택 범위가 데이터 영역 안이고 전체가 아니면 {top,bottom} 반환, 아니면 null */
function getSelectionScope(sh, lastRow) {
  var rng = sh.getActiveRange();
  if (!rng) return null;
  var top = rng.getRow(), bottom = top + rng.getNumRows() - 1;
  if (top < CONFIG.DATA_START_ROW) return null;          // 헤더 포함 선택 = 전체로 간주
  if (top <= CONFIG.DATA_START_ROW && bottom >= lastRow) return null; // 전체 선택
  if (rng.getNumRows() <= 1) return null;                // 한 칸만 클릭 = 전체로 간주
  return { top: top, bottom: bottom };
}

/** '2025.02.' / '2025-02' / Date → 'YYYYMM' */
function parseYm(v) {
  if (v === null || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return v.getFullYear() + ('0' + (v.getMonth() + 1)).slice(-2);
  }
  var m = String(v).match(/(\d{4})\D*(\d{1,2})/);
  if (!m) return null;
  return m[1] + ('0' + m[2]).slice(-2);
}

/** 'YYYYMM' → 그 달 말일 'YYYYMMDD' */
function lastDayYmd(ym) {
  var y = parseInt(ym.substr(0, 4), 10);
  var mo = parseInt(ym.substr(4, 2), 10);
  var d = new Date(y, mo, 0).getDate();  // mo월 0일 = mo-1월 말일 계산 트릭
  return ym + ('0' + d).slice(-2);
}

/** 숫자/문자(콤마 포함) → 정수. 실패 시 null */
function toInt(v) {
  if (v === null || v === '') return null;
  var n = Number(String(v).replace(/[, ]/g, ''));
  if (isNaN(n)) return null;
  return Math.round(n);
}

/** 주민번호 정규화: 숫자만 추출 */
function normalizeRrn(v) {
  if (v === null || v === undefined) return '';
  var d = String(v).replace(/[^0-9]/g, '');
  return d;
}

function cell(row, colIdx) { return colIdx > 0 ? row[colIdx - 1] : ''; }
function notEmpty(v) { return v !== null && v !== undefined && String(v).trim() !== ''; }
function objValues(o) { return Object.keys(o).map(function (k) { return o[k]; }); }
function uniq(a) { var s = {}, r = []; a.forEach(function (x) { if (!s[x]) { s[x] = 1; r.push(x); } }); return r; }
function stamp() {
  var d = new Date();
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyyMMdd_HHmm');
}
