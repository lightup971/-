/* ============================================================================
 *  브라우저에서 .xlsx 읽기 — 외부 라이브러리 없음
 * ----------------------------------------------------------------------------
 *  품의 첨부파일을 열지 않고 그대로 업로드해서 파싱하기 위한 최소 구현.
 *  · ZIP(중앙 디렉터리) 파싱 → stored/deflate 해제(DecompressionStream)
 *  · sharedStrings / worksheet XML 파싱 → 2차원 배열
 *  · 병합셀은 확장하지 않는다(엑셀에서 복사해 붙여넣은 것과 동일한 모양 유지).
 *    품의 표는 소득자명이 첫 행에만 있는 구조라 이 모양이 오히려 정확하다.
 * ==========================================================================*/
(function (root) {
  'use strict';

  function u16(b, p) { return b[p] | (b[p + 1] << 8); }
  function u32(b, p) { return (b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24)) >>> 0; }

  /** ZIP 엔트리 목록 { name, offset, compMethod, compSize, size } */
  function listEntries(bytes) {
    // End of Central Directory 찾기 (뒤에서부터)
    var eocd = -1;
    for (var i = bytes.length - 22; i >= 0 && i > bytes.length - 66000; i--) {
      if (u32(bytes, i) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('ZIP 구조를 찾지 못했습니다(올바른 .xlsx 파일이 아닐 수 있습니다).');
    var count = u16(bytes, eocd + 10);
    var cdOff = u32(bytes, eocd + 16);

    var out = [], p = cdOff;
    for (var n = 0; n < count; n++) {
      if (u32(bytes, p) !== 0x02014b50) break;
      var method = u16(bytes, p + 10);
      var compSize = u32(bytes, p + 20);
      var size = u32(bytes, p + 24);
      var nameLen = u16(bytes, p + 28);
      var extraLen = u16(bytes, p + 30);
      var cmtLen = u16(bytes, p + 32);
      var localOff = u32(bytes, p + 42);
      var name = new TextDecoder('utf-8').decode(bytes.subarray(p + 46, p + 46 + nameLen));
      out.push({ name: name, method: method, compSize: compSize, size: size, localOff: localOff });
      p += 46 + nameLen + extraLen + cmtLen;
    }
    return out;
  }

  async function readEntry(bytes, e) {
    // 로컬 헤더에서 실제 데이터 시작 위치 계산
    var p = e.localOff;
    if (u32(bytes, p) !== 0x04034b50) throw new Error('ZIP 항목이 손상되었습니다: ' + e.name);
    var nameLen = u16(bytes, p + 26), extraLen = u16(bytes, p + 28);
    var start = p + 30 + nameLen + extraLen;
    var raw = bytes.subarray(start, start + e.compSize);
    if (e.method === 0) return new TextDecoder('utf-8').decode(raw);
    if (e.method !== 8) throw new Error('지원하지 않는 압축 방식입니다(' + e.method + ').');
    if (typeof DecompressionStream === 'undefined') {
      throw new Error('이 브라우저는 파일 해제를 지원하지 않습니다. 최신 Chrome/Edge에서 열어 주세요.');
    }
    var ds = new DecompressionStream('deflate-raw');
    var stream = new Blob([raw]).stream().pipeThrough(ds);
    var buf = await new Response(stream).arrayBuffer();
    return new TextDecoder('utf-8').decode(new Uint8Array(buf));
  }

  function parseXml(text) {
    var doc = new DOMParser().parseFromString(text, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('엑셀 내부 XML을 읽지 못했습니다.');
    return doc;
  }

  /** A1 → {col:0-based, row:0-based} */
  function refToRC(ref) {
    var m = /^([A-Z]+)(\d+)$/.exec(ref || '');
    if (!m) return null;
    var c = 0;
    for (var i = 0; i < m[1].length; i++) c = c * 26 + (m[1].charCodeAt(i) - 64);
    return { col: c - 1, row: parseInt(m[2], 10) - 1 };
  }

  function textOf(node) {
    // <si> 안의 모든 <t> 이어붙이기 (리치텍스트 대응)
    var ts = node.getElementsByTagName('t'), s = '';
    for (var i = 0; i < ts.length; i++) s += ts[i].textContent;
    return s;
  }

  /** .xlsx ArrayBuffer → [{name, rows:[[cell,…],…]}] */
  async function read(arrayBuffer) {
    var bytes = new Uint8Array(arrayBuffer);
    var entries = listEntries(bytes);
    var byName = {};
    entries.forEach(function (e) { byName[e.name] = e; });

    function need(n) {
      if (!byName[n]) throw new Error('엑셀 구성요소를 찾지 못했습니다: ' + n);
      return byName[n];
    }

    // 공유 문자열
    var shared = [];
    if (byName['xl/sharedStrings.xml']) {
      var sdoc = parseXml(await readEntry(bytes, byName['xl/sharedStrings.xml']));
      var sis = sdoc.getElementsByTagName('si');
      for (var i = 0; i < sis.length; i++) shared.push(textOf(sis[i]));
    }

    // 시트 목록(순서/이름) + rId → 파일경로
    var wdoc = parseXml(await readEntry(bytes, need('xl/workbook.xml')));
    var rdoc = parseXml(await readEntry(bytes, need('xl/_rels/workbook.xml.rels')));
    var relMap = {};
    var rels = rdoc.getElementsByTagName('Relationship');
    for (var r = 0; r < rels.length; r++) {
      var t = rels[r].getAttribute('Target') || '';
      relMap[rels[r].getAttribute('Id')] = t.charAt(0) === '/' ? t.slice(1) : ('xl/' + t.replace(/^\.\//, ''));
    }

    var sheets = [];
    var sh = wdoc.getElementsByTagName('sheet');
    for (var s = 0; s < sh.length; s++) {
      var name = sh[s].getAttribute('name');
      // state="hidden" / "veryHidden" 은 엑셀 화면에 보이지 않는 시트
      var state = sh[s].getAttribute('state') || 'visible';
      var rid = sh[s].getAttribute('r:id') || sh[s].getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
      var target = relMap[rid];
      if (!target || !byName[target]) continue;
      sheets.push({
        name: name,
        hidden: state !== 'visible',
        rows: await readSheet(bytes, byName[target], shared)
      });
    }
    return sheets;
  }

  async function readSheet(bytes, entry, shared) {
    var doc = parseXml(await readEntry(bytes, entry));
    var rowsEl = doc.getElementsByTagName('row');
    var grid = [], maxCol = 0;

    for (var i = 0; i < rowsEl.length; i++) {
      var rEl = rowsEl[i];
      var rIdx = parseInt(rEl.getAttribute('r') || (i + 1), 10) - 1;
      var cells = rEl.getElementsByTagName('c');
      var line = grid[rIdx] || (grid[rIdx] = []);
      for (var j = 0; j < cells.length; j++) {
        var c = cells[j];
        var rc = refToRC(c.getAttribute('r'));
        var col = rc ? rc.col : j;
        var t = c.getAttribute('t');
        var val = '';
        if (t === 'inlineStr') {
          val = textOf(c);
        } else {
          var vs = c.getElementsByTagName('v');
          var raw = vs.length ? vs[0].textContent : '';
          if (t === 's') val = shared[parseInt(raw, 10)] || '';
          else if (t === 'str' || t === 'e') val = raw;
          else val = raw;              // 숫자/날짜 시리얼은 원문 그대로
        }
        line[col] = val;
        if (col + 1 > maxCol) maxCol = col + 1;
      }
    }
    // 구멍 메우기
    var out = [];
    for (var k = 0; k < grid.length; k++) {
      var row = grid[k] || [];
      var filled = new Array(maxCol);
      for (var m = 0; m < maxCol; m++) filled[m] = row[m] == null ? '' : String(row[m]);
      out.push(filled);
    }
    return out;
  }

  var API = { read: read, listEntries: listEntries };
  if (typeof module !== 'undefined' && module.exports) module.exports = API;
  else root.XlsxReader = API;
})(typeof self !== 'undefined' ? self : this);
