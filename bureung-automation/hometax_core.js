/* ============================================================================
 *  홈택스 간이지급명세서(거주자의 사업소득) 업로드 엑셀 생성 — 순수 JS 코어
 * ----------------------------------------------------------------------------
 *  국세청 양식(Sheet1) 구조 — 컬럼마다 타입이 다르므로 그대로 재현해야 함:
 *    A 일련번호        숫자(General)  1부터 순차
 *    B 귀속연도        숫자(General)  예: 2024
 *    C 귀속월          문자(@)        예: '01'  ← 앞자리 0 유지
 *    D 업종코드        숫자(@서식)    940918(퀵서비스)
 *    E 소득자성명(상호) 문자(@)
 *    F 주민(사업자)등록번호 문자(@)    13자리, 하이픈 없음
 *    G 내외국인        문자(@)        '1'=내국인 / '9'=외국인
 *    H 지급액          숫자(#,##0)
 *    I 세율            문자(@)        '3'
 *    J 소득세          숫자(#,##0)
 *    K 지방소득세      숫자(#,##0)
 *
 *  제출요령(원본):
 *   · 반드시 2행부터 데이터 입력, 1,000라인 초과 불가
 *   · Sheet명 수정 시 오류 → 'Sheet1' 유지
 *   · 지급액 음수 불가, 없으면 0
 *   · 입력한 행 이후에 임의의 값이 있으면 오류
 *
 *  월별로 파일을 분리 생성한다(귀속월 = 구글시트 지급연월).
 * ==========================================================================*/
(function (root) {
  'use strict';

  var HT_HEADER = [
    '일련번호\n1부터 순차로 기재', '귀속연도', "귀속월\n예시) 1월인 경우  '01' ", '업종코드',
    '소득자성명(상호)', '주민(사업자)등록번호', "내외국인\n'1' 또는 '9' 기재",
    '지급액', '세율', '소득세', '지방소득세'
  ];

  var HT_DEFAULTS = {
    BIZTP_CD: 940918,     // 업종코드(퀵서비스) — 숫자
    TAX_RT: '3',          // 세율 — 문자
    MERGE_DUP: true,      // 같은 주민번호+귀속월 합산
    MAX_ROWS: 1000,       // 홈택스 1파일 최대 라인
    SHEET_NAME: 'Sheet1'
  };

  /* 주민번호 뒷자리 첫 숫자로 내/외국인 판별: 1~4 내국인('1'), 5~8 외국인('9') */
  function foreignFlag(rrn13) {
    var d = String(rrn13 || '').replace(/[^0-9]/g, '');
    if (d.length !== 13) return { flag: '1', ok: false };
    var g = d.charAt(6);
    if ('1234'.indexOf(g) >= 0) return { flag: '1', ok: true };
    if ('5678'.indexOf(g) >= 0) return { flag: '9', ok: true };
    return { flag: '1', ok: false };
  }

  /**
   * ERP 코어의 buildRows() 결과(rows: {kind,name,code,rrn,ym,amt,intax,locint})를
   * 월별 홈택스 파일 데이터로 변환.
   * @return {files:[{ym, year, month, rows:[...], name}], warnings:{...}}
   */
  function buildHometaxFiles(built, opt) {
    opt = opt || {};
    var cfg = Object.assign({}, HT_DEFAULTS, opt);
    var src = built.rows || [];
    var warn = { noRrn: [], badFlag: [], merged: 0, overflow: [], months: 0, total: 0 };

    // 1) 귀속월별로 묶기
    var byYm = {}, order = [];
    src.forEach(function (r) {
      if (!r.ym) return;
      if (!byYm[r.ym]) { byYm[r.ym] = []; order.push(r.ym); }
      byYm[r.ym].push(r);
    });
    order.sort();

    var files = order.map(function (ym) {
      var list = byYm[ym];

      // 2) 같은 주민번호는 합산 (홈택스는 소득자별 1행)
      var out = [], idx = {};
      list.forEach(function (r) {
        var rrn = String(r.rrn || '').replace(/[^0-9]/g, '');
        var key = rrn || ('@name:' + r.name + '/' + r.kind);   // 주민번호 없으면 합치지 않음
        if (cfg.MERGE_DUP && rrn && idx[key] !== undefined) {
          var t = out[idx[key]];
          t.amt += r.amt; t.intax += r.intax; t.locint += r.locint;
          t.mergedCount++;
          warn.merged++;
          return;
        }
        var ff = foreignFlag(rrn);
        if (!rrn) {
          if (warn.noRrn.indexOf(r.name) < 0) warn.noRrn.push(r.name || '(이름없음)');
        } else if (!ff.ok) {
          if (warn.badFlag.indexOf(r.name) < 0) warn.badFlag.push(r.name || '(이름없음)');
        }
        idx[key] = out.length;
        out.push({
          name: r.name, rrn: rrn, kind: r.kind, foreign: ff.flag,
          amt: r.amt, intax: r.intax, locint: r.locint, mergedCount: 1
        });
      });

      if (out.length > cfg.MAX_ROWS) warn.overflow.push(ym);
      warn.total += out.length;
      return {
        ym: ym, year: parseInt(ym.substr(0, 4), 10), month: ym.substr(4, 2),
        rows: out, fileName: '간이지급명세서_' + ym + '.xlsx'
      };
    });
    warn.months = files.length;
    return { files: files, warnings: warn, cfg: cfg };
  }

  /** 한 달치 파일 → Sheet1 셀 배열(AOA). 각 셀은 {v, t:'n'|'s', s:styleIdx} */
  function toSheetCells(file, cfg) {
    var S_GEN = 0, S_TXT = 1, S_NUM = 2;
    var aoa = [HT_HEADER.map(function (h) { return { v: h, t: 's', s: S_TXT }; })];
    file.rows.forEach(function (r, i) {
      aoa.push([
        { v: i + 1,             t: 'n', s: S_GEN },   // A 일련번호
        { v: file.year,         t: 'n', s: S_GEN },   // B 귀속연도
        { v: file.month,        t: 's', s: S_TXT },   // C 귀속월 '01'
        { v: cfg.BIZTP_CD,      t: 'n', s: S_TXT },   // D 업종코드(값 숫자, 서식 @)
        { v: r.name,            t: 's', s: S_TXT },   // E 성명
        { v: r.rrn,             t: 's', s: S_TXT },   // F 주민번호 13자리
        { v: r.foreign,         t: 's', s: S_TXT },   // G 내외국인
        { v: Math.max(0, r.amt), t: 'n', s: S_NUM },  // H 지급액(음수 불가)
        { v: cfg.TAX_RT,        t: 's', s: S_TXT },   // I 세율
        { v: Math.max(0, r.intax),  t: 'n', s: S_NUM },  // J 소득세
        { v: Math.max(0, r.locint), t: 'n', s: S_NUM }   // K 지방소득세
      ]);
    });
    return aoa;
  }

  /* ---------------- XLSX (타입 혼합 지원) ---------------- */
  function xmlEsc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  function colLetter(n) { var s = ''; n++; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }

  function sheetXmlTyped(aoa) {
    var rows = aoa.map(function (arr, ri) {
      var cells = arr.map(function (c, ci) {
        if (c == null || c.v === '' || c.v == null) return '';
        var ref = colLetter(ci) + (ri + 1);
        var s = ' s="' + (c.s || 0) + '"';
        if (c.t === 'n') return '<c r="' + ref + '"' + s + '><v>' + c.v + '</v></c>';
        return '<c r="' + ref + '"' + s + ' t="inlineStr"><is><t xml:space="preserve">' + xmlEsc(c.v) + '</t></is></c>';
      }).join('');
      return '<row r="' + (ri + 1) + '">' + cells + '</row>';
    }).join('');
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<cols><col min="1" max="1" width="15.75" customWidth="1"/><col min="2" max="2" width="11.8" customWidth="1"/>' +
      '<col min="3" max="3" width="20.25" customWidth="1"/><col min="4" max="4" width="9" customWidth="1"/>' +
      '<col min="5" max="5" width="37.25" customWidth="1"/><col min="6" max="6" width="19.58" customWidth="1"/>' +
      '<col min="7" max="7" width="14.58" customWidth="1"/><col min="8" max="8" width="19.25" customWidth="1"/>' +
      '<col min="9" max="9" width="8" customWidth="1"/><col min="10" max="11" width="19.25" customWidth="1"/></cols>' +
      '<sheetData>' + rows + '</sheetData></worksheet>';
  }

  var GUIDE_TEXT =
    '   ※ 반드시 2행부터 데이터를 입력하여야 하며, 1,000라인을 초과하여 작성할 수 없습니다.\n' +
    '   ※ 파일명은 변경이 가능하나 30자를 초과할 수 없습니다.\n' +
    '   ※ Sheet명을 수정하는 경우 오류가 발생할 수 있으니 수정하지 마시기 바랍니다.\n' +
    '   ※ 사용방법\n' +
    '      1. 찾아보기 버튼을 눌러 작성한 간이지급명세서 엑셀파일을 선택한 후, 업로드합니다.\n' +
    '      2. 검증하기 버튼을 눌러 형식 및 내용을 검증합니다. 검증결과 확인 버튼으로 결과를 확인합니다.\n' +
    '      3. 오류가 발생한 경우 오류내역을 확인하여 엑셀파일을 정정한 후 1번부터 다시 수행합니다.\n' +
    '      4. 오류가 없는 경우 창을 닫고 작성목록에서 내용을 확인합니다.\n' +
    '      6. 과세자료 작성완료 버튼을 눌러 간이지급명세서를 제출하고 접수증을 출력합니다.\n' +
    '   ※ 주의사항\n' +
    '      1. 소득자의 간이지급명세를 입력한 행 이외에는 공란을 유지하여야 합니다.\n' +
    '      2. 지급액 항목은 음수를 입력할 수 없으며, 지급한 금액이 없는 경우에는 0을 입력합니다.';

  var BIZ_CODES = [[940100,'저술가'],[940200,'화가관련'],[940301,'작곡가'],[940302,'배우'],[940303,'모델'],[940304,'가수'],[940305,'성악가'],[940306,'1인미디어 콘텐츠창작자'],[940500,'연예보조'],[940600,'자문ㆍ고문'],[940901,'바둑기사'],[940902,'꽃꽃이교사'],[940903,'학원강사'],[940904,'직업운동가'],[940905,'봉사료수취자'],[940906,'보험설계'],[940907,'음료배달'],[940908,'방판. 외판'],[940909,'기타자영업'],[940910,'다단계판매'],[940911,'기타모집수당'],[940912,'간병인'],[940913,'대리운전'],[940914,'캐디'],[940915,'목욕관리사'],[940916,'행사도우미'],[940917,'심부름용역'],[940918,'퀵서비스'],[940919,'물품배달'],[851101,'병의원'],[940920,'학습지방문강사'],[940921,'교육교구방문강사'],[940922,'대여제품방문점검원'],[940923,'대출모집인'],[940924,'신용카드회원모집인'],[940925,'방과후강사'],[940926,'소프트웨어프리랜서'],[940927,'관광통역안내사'],[940928,'어린이통학버스기사'],[940929,'중고자동차판매원']];

  var HT_STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="2"><numFmt numFmtId="164" formatCode="@"/>' +
    '<numFmt numFmtId="165" formatCode="#,##0_);[Red]\\(#,##0\\)"/></numFmts>' +
    '<fonts count="1"><font><sz val="11"/><name val="맑은 고딕"/></font></fonts>' +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border/></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="3">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

  /* ZIP (stored) */
  var _crc = (function () { var t = []; for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); t[n] = c >>> 0; } return t; })();
  function crc32(b) { var c = 0xFFFFFFFF; for (var i = 0; i < b.length; i++) c = _crc[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  function utf8(s) { return new TextEncoder().encode(s); }
  function u16(a, v) { a.push(v & 0xFF, (v >>> 8) & 0xFF); }
  function u32(a, v) { a.push(v & 0xFF, (v >>> 8) & 0xFF, (v >>> 16) & 0xFF, (v >>> 24) & 0xFF); }
  function zip(files) {
    var local = [], central = [], offset = 0;
    files.forEach(function (f) {
      var nameB = utf8(f.name), crc = crc32(f.data), sz = f.data.length;
      var lh = []; u32(lh, 0x04034b50); u16(lh, 20); u16(lh, 0x0800); u16(lh, 0); u16(lh, 0); u16(lh, 0);
      u32(lh, crc); u32(lh, sz); u32(lh, sz); u16(lh, nameB.length); u16(lh, 0);
      var lhb = new Uint8Array(lh.length + nameB.length + sz);
      lhb.set(lh, 0); lhb.set(nameB, lh.length); lhb.set(f.data, lh.length + nameB.length);
      local.push(lhb);
      var ch = []; u32(ch, 0x02014b50); u16(ch, 20); u16(ch, 20); u16(ch, 0x0800); u16(ch, 0); u16(ch, 0); u16(ch, 0);
      u32(ch, crc); u32(ch, sz); u32(ch, sz);
      u16(ch, nameB.length); u16(ch, 0); u16(ch, 0); u16(ch, 0); u16(ch, 0); u32(ch, 0); u32(ch, offset);
      var chb = new Uint8Array(ch.length + nameB.length);
      chb.set(ch, 0); chb.set(nameB, ch.length);
      central.push(chb); offset += lhb.length;
    });
    var cs = central.reduce(function (s, b) { return s + b.length; }, 0);
    var e = []; u32(e, 0x06054b50); u16(e, 0); u16(e, 0); u16(e, files.length); u16(e, files.length);
    u32(e, cs); u32(e, offset); u16(e, 0);
    var out = new Uint8Array(offset + cs + e.length), p = 0;
    local.forEach(function (b) { out.set(b, p); p += b.length; });
    central.forEach(function (b) { out.set(b, p); p += b.length; });
    out.set(new Uint8Array(e), p);
    return out;
  }

  /** 한 달치 홈택스 xlsx 생성 (Sheet1 + 제출요령 + 사업소득업종코드) */
  function buildHometaxXlsx(file, cfg) {
    cfg = Object.assign({}, HT_DEFAULTS, cfg || {});
    var s1 = sheetXmlTyped(toSheetCells(file, cfg));

    var guide = [[{ v: GUIDE_TEXT, t: 's', s: 1 }]];
    var codes = [[{ v: '사업소득업종코드', t: 's', s: 1 }, { v: '사업소득업종명', t: 's', s: 1 }]]
      .concat(BIZ_CODES.map(function (c) { return [{ v: c[0], t: 'n', s: 0 }, { v: c[1], t: 's', s: 1 }]; }));

    var parts = [
      { name: '[Content_Types].xml', data: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '</Types>') },
      { name: '_rels/.rels', data: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>') },
      { name: 'xl/workbook.xml', data: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="' + xmlEsc(cfg.SHEET_NAME) + '" sheetId="1" r:id="rId1"/>' +
        '<sheet name=" 제출요령" sheetId="2" r:id="rId2"/>' +
        '<sheet name="사업소득업종코드" sheetId="3" r:id="rId3"/></sheets></workbook>') },
      { name: 'xl/_rels/workbook.xml.rels', data: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' +
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>' +
        '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '</Relationships>') },
      { name: 'xl/styles.xml', data: utf8(HT_STYLES) },
      { name: 'xl/worksheets/sheet1.xml', data: utf8(s1) },
      { name: 'xl/worksheets/sheet2.xml', data: utf8(sheetXmlTyped(guide)) },
      { name: 'xl/worksheets/sheet3.xml', data: utf8(sheetXmlTyped(codes)) }
    ];
    return zip(parts);
  }

  /** 여러 월 파일을 하나의 ZIP으로 */
  function buildZipOfFiles(files, cfg) {
    var entries = files.map(function (f) {
      return { name: f.fileName, data: buildHometaxXlsx(f, cfg) };
    });
    return zip(entries);
  }

  var API = {
    buildHometaxFiles: buildHometaxFiles, buildHometaxXlsx: buildHometaxXlsx,
    buildZipOfFiles: buildZipOfFiles, toSheetCells: toSheetCells,
    foreignFlag: foreignFlag, HT_HEADER: HT_HEADER, HT_DEFAULTS: HT_DEFAULTS
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = API;
  else root.HometaxCore = API;
})(typeof self !== 'undefined' ? self : this);
