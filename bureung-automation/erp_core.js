/* ============================================================================
 *  ERP 업로드 엑셀 생성 — 순수 JS 코어 (외부 라이브러리 0개)
 *  - parseClipboard : 구글시트 복사본(TSV, 따옴표/줄바꿈 지원) 파싱
 *  - buildRows      : H~AC 블록 → ERP 11컬럼 데이터 + 경고
 *  - buildXlsx      : 데이터 → 유효한 .xlsx (Uint8Array), 전 셀 텍스트서식
 *  브라우저/Node 공용 (TextEncoder, Uint8Array 사용)
 * ==========================================================================*/
(function (root) {
  'use strict';

  var ERP_CODES  = ['INCMPER_NO','PAY_YM','RVERS_YM','PAY_DT','BIZAREA_CD','DEPT_CD','WGS_AMT','BIZTP_FG_CD','TAX_RT','INTAX_AMT','LOCINTAX_AMT'];
  var ERP_LABELS = ['소득자코드','지급연월','귀속년월','지급일','사업장코드','부서코드','지급액','업종구분코드','세율','소득세금액','지방소득세금액'];

  var DEFAULTS = {
    COMPANY_CD: '1000', COMPANY_NM: '(주)부릉', BIZR_NO: '2068673707',
    BIZAREA_CD: '1000', DEPT_CD: 'AE0000000', BIZTP_FG_CD: '940918', TAX_RT: '3',
    CODE_PAD_LEN: 6, AMOUNT_BASIS: 'FINAL', SHEET_NAME: 'TSMINC00700_F'
  };

  // 붙여넣는 블록은 H열부터 AC열까지 → 0-based 오프셋
  var OFF = {
    부인자명: 1, 부인주민: 2, 지급연월: 4, 부인소득: 5, 부인소득세: 8, 부인지방: 9,
    해당자명: 14, 해당주민: 15, 기존소득: 16, 기존소득세: 17, 기존지방: 18,
    최종소득: 19, 최종소득세: 20, 최종지방: 21
  };
  var BLOCK_WIDTH = 22; // H..AC

  /* ---------- 파서 ---------- */
  function parseClipboard(text) {
    var rows = [], row = [], field = '', i = 0, inQ = false;
    text = String(text).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    while (i < text.length) {
      var ch = text[i];
      if (inQ) {
        if (ch === '"') {
          if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
          inQ = false; i++; continue;
        }
        field += ch; i++; continue;
      }
      if (ch === '"') { inQ = true; i++; continue; }
      if (ch === '\t') { row.push(field); field = ''; i++; continue; }
      if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; i++; continue; }
      field += ch; i++;
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    // 완전히 빈 행 제거
    return rows.filter(function (r) { return r.some(function (c) { return String(c).trim() !== ''; }); });
  }

  /* ---------- 유틸 ---------- */
  function parseNameCode(v) {
    if (v == null) return { name: '', code: '' };
    var s = String(v).trim();
    var nums = s.match(/\d+/g);
    var code = nums ? nums[nums.length - 1] : '';
    var name = code ? s.replace(new RegExp('[\\s\\n]*' + code + '[\\s\\n]*$'), '') : s;
    return { name: name.replace(/[\s\n]+/g, ' ').trim(), code: code };
  }
  function padCode(code, n) {
    if (!code) return '';
    var c = String(code).trim();
    if (n > 0 && /^\d+$/.test(c)) while (c.length < n) c = '0' + c;
    return c;
  }
  function parseYm(v) {
    if (v == null || v === '') return null;
    var m = String(v).match(/(\d{4})\D*(\d{1,2})/);
    if (!m) return null;
    return m[1] + ('0' + m[2]).slice(-2);
  }
  function lastDayYmd(ym) {
    var y = parseInt(ym.substr(0, 4), 10), mo = parseInt(ym.substr(4, 2), 10);
    var d = new Date(y, mo, 0).getDate();
    return ym + ('0' + d).slice(-2);
  }
  function toInt(v) {
    if (v == null || v === '') return null;
    var n = Number(String(v).replace(/[, ]/g, ''));
    return isNaN(n) ? null : Math.round(n);
  }
  function cell(r, off) { return off < r.length ? r[off] : ''; }
  function notEmpty(v) { return v != null && String(v).trim() !== ''; }

  /* ---------- 변환: 붙여넣은 블록 → ERP 데이터 ---------- */
  function buildRows(matrix, opt) {
    opt = opt || {};
    var cfg = Object.assign({}, DEFAULTS, opt);
    var out = [], warnings = { missingCode: [], fallback: 0, badBlockWidth: false, orphan: 0 };
    var current = null;

    matrix.forEach(function (r) {
      if (r.length && r.length < 15) warnings.badBlockWidth = true; // H~AC 미만이면 열 어긋남 의심

      if (notEmpty(cell(r, OFF.부인자명))) {
        var parsed = parseNameCode(cell(r, OFF.해당자명));
        current = { name: parsed.name, code: padCode(parsed.code, cfg.CODE_PAD_LEN) };
      }
      var ym = parseYm(cell(r, OFF.지급연월));
      if (ym === null) return;
      if (!current) { warnings.orphan++; return; }

      var amt, intax, locint;
      if (cfg.AMOUNT_BASIS === 'DENIED') {
        amt = toInt(cell(r, OFF.부인소득));
        intax = toInt(cell(r, OFF.부인소득세)) || 0;
        locint = toInt(cell(r, OFF.부인지방)) || 0;
      } else {
        amt = toInt(cell(r, OFF.최종소득));
        intax = toInt(cell(r, OFF.최종소득세));
        locint = toInt(cell(r, OFF.최종지방));
        if (amt === null) { // 최종 공란 → 기존 + 부인
          amt = (toInt(cell(r, OFF.기존소득)) || 0) + (toInt(cell(r, OFF.부인소득)) || 0);
          intax = (toInt(cell(r, OFF.기존소득세)) || 0) + (toInt(cell(r, OFF.부인소득세)) || 0);
          locint = (toInt(cell(r, OFF.기존지방)) || 0) + (toInt(cell(r, OFF.부인지방)) || 0);
          if (amt > 0) warnings.fallback++;
        }
        intax = intax || 0; locint = locint || 0;
      }
      if (amt === null || amt <= 0) return;

      var missing = !current.code;
      if (missing && warnings.missingCode.indexOf(current.name) < 0) warnings.missingCode.push(current.name || '(이름없음)');

      out.push({
        cells: [current.code || '', ym, ym, lastDayYmd(ym), cfg.BIZAREA_CD, cfg.DEPT_CD,
                String(amt), cfg.BIZTP_FG_CD, cfg.TAX_RT, String(intax), String(locint)],
        name: current.name, missing: missing
      });
    });
    return { rows: out, warnings: warnings, cfg: cfg };
  }

  /* ---------- XLSX 생성 (stored-zip + inlineStr 텍스트) ---------- */
  function xmlEsc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  function colLetter(n) { // 0-based → A,B,...
    var s = ''; n++;
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
    return s;
  }
  function sheetXml(aoa) {
    var rows = aoa.map(function (arr, ri) {
      var cells = arr.map(function (val, ci) {
        var ref = colLetter(ci) + (ri + 1);
        if (val === '' || val == null) return '';
        return '<c r="' + ref + '" t="inlineStr" s="1"><is><t xml:space="preserve">' + xmlEsc(val) + '</t></is></c>';
      }).join('');
      return '<row r="' + (ri + 1) + '">' + cells + '</row>';
    }).join('');
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
      rows + '</sheetData></worksheet>';
  }

  function buildAoa(built) {
    var cfg = built.cfg;
    var aoa = [
      ['COMPANY_CD', 'COMPANY_NM', 'BIZR_NO'],
      ['회사코드', '회사명', '사업자번호'],
      [cfg.COMPANY_CD, cfg.COMPANY_NM, cfg.BIZR_NO],
      [],
      ERP_CODES.slice(),
      ERP_LABELS.slice()
    ];
    built.rows.forEach(function (r) { aoa.push(r.cells); });
    return aoa;
  }

  // ---- ZIP (stored) ----
  var _crcTable = (function () {
    var t = [];
    for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); t[n] = c >>> 0; }
    return t;
  })();
  function crc32(bytes) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) c = _crcTable[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }
  function utf8(str) { return new TextEncoder().encode(str); }
  function pushU16(a, v) { a.push(v & 0xFF, (v >>> 8) & 0xFF); }
  function pushU32(a, v) { a.push(v & 0xFF, (v >>> 8) & 0xFF, (v >>> 16) & 0xFF, (v >>> 24) & 0xFF); }

  function zip(files) {
    // files: [{name, data:Uint8Array}]
    var local = [], central = [], offset = 0;
    files.forEach(function (f) {
      var nameB = utf8(f.name), crc = crc32(f.data), sz = f.data.length;
      var lh = [];
      pushU32(lh, 0x04034b50); pushU16(lh, 20); pushU16(lh, 0); pushU16(lh, 0); pushU16(lh, 0); pushU16(lh, 0);
      pushU32(lh, crc); pushU32(lh, sz); pushU32(lh, sz); pushU16(lh, nameB.length); pushU16(lh, 0);
      var lhb = new Uint8Array(lh.length + nameB.length + sz);
      lhb.set(lh, 0); lhb.set(nameB, lh.length); lhb.set(f.data, lh.length + nameB.length);
      local.push(lhb);

      var ch = [];
      pushU32(ch, 0x02014b50); pushU16(ch, 20); pushU16(ch, 20); pushU16(ch, 0); pushU16(ch, 0); pushU16(ch, 0); pushU16(ch, 0);
      pushU32(ch, crc); pushU32(ch, sz); pushU32(ch, sz);
      pushU16(ch, nameB.length); pushU16(ch, 0); pushU16(ch, 0); pushU16(ch, 0); pushU16(ch, 0); pushU32(ch, 0);
      pushU32(ch, offset);
      var chb = new Uint8Array(ch.length + nameB.length);
      chb.set(ch, 0); chb.set(nameB, ch.length);
      central.push(chb);
      offset += lhb.length;
    });
    var centralSize = central.reduce(function (s, b) { return s + b.length; }, 0);
    var eocd = [];
    pushU32(eocd, 0x06054b50); pushU16(eocd, 0); pushU16(eocd, 0);
    pushU16(eocd, files.length); pushU16(eocd, files.length);
    pushU32(eocd, centralSize); pushU32(eocd, offset); pushU16(eocd, 0);

    var total = offset + centralSize + eocd.length;
    var out = new Uint8Array(total), p = 0;
    local.forEach(function (b) { out.set(b, p); p += b.length; });
    central.forEach(function (b) { out.set(b, p); p += b.length; });
    out.set(new Uint8Array(eocd), p);
    return out;
  }

  var STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="@"/></numFmts>' +
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border/></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

  function buildXlsx(built) {
    var sheetName = xmlEsc(built.cfg.SHEET_NAME);
    var files = [
      { name: '[Content_Types].xml', data: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
          '<Default Extension="xml" ContentType="application/xml"/>' +
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
          '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
          '</Types>') },
      { name: '_rels/.rels', data: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
          '</Relationships>') },
      { name: 'xl/workbook.xml', data: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
          '<sheets><sheet name="' + sheetName + '" sheetId="1" r:id="rId1"/></sheets></workbook>') },
      { name: 'xl/_rels/workbook.xml.rels', data: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
          '</Relationships>') },
      { name: 'xl/styles.xml', data: utf8(STYLES) },
      { name: 'xl/worksheets/sheet1.xml', data: utf8(sheetXml(buildAoa(built))) }
    ];
    return zip(files);
  }

  var API = { parseClipboard: parseClipboard, buildRows: buildRows, buildXlsx: buildXlsx,
              buildAoa: buildAoa, ERP_CODES: ERP_CODES, ERP_LABELS: ERP_LABELS, DEFAULTS: DEFAULTS };
  if (typeof module !== 'undefined' && module.exports) module.exports = API;
  else root.ERPCore = API;
})(typeof self !== 'undefined' ? self : this);
