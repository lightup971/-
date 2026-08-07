/* ============================================================================
 *  품의(업무협조 기안 첨부 xlsx) → 구글시트 붙여넣기 블록 변환 코어
 * ----------------------------------------------------------------------------
 *  매뉴얼 [00. 재무회계실 수신] + [01. ERP 데이터 수정] 앞단을 줄인다.
 *
 *  문제: 품의 xlsx 컬럼과 구글시트 컬럼이 1:1이 아니다.
 *    · 품의는 시트마다 배치가 다름 (기사코드 열이 있는 것/없는 것)
 *    · 시트에는 S·T·U(부인자 최종)가 중간에 끼어 있어 한 번에 붙여넣을 수 없음
 *    · 소득자코드는 ERP에서 한 명씩 조회해 이름 뒤에 Alt+Enter로 붙여야 함
 *
 *  해결: 헤더 이름으로 열을 인식해 시트 H~AC(22열) 배치로 재구성하고,
 *        소득자코드는 과거 이력 사전에서 자동으로 채운다.
 * ==========================================================================*/
(function (root) {
  'use strict';

  /* 구글시트 H~AC (0-based). ERP/홈택스 생성기와 동일한 배치 */
  var OUT = {
    일자: 0, 소득자명: 1, 부인주민: 2, 기사코드: 3, 지급연월: 4, 소득금액: 5,
    세율: 6, 실지급액: 7, 소득세: 8, 지방소득세: 9, 부인신청내역: 10,
    부인최종소득: 11, 부인최종소득세: 12, 부인최종지방: 13,   // S,T,U
    해당자명: 14, 해당주민: 15,
    해당기존소득: 16, 해당기존소득세: 17, 해당기존지방: 18,   // X,Y,Z (ERP 조회)
    최종소득: 19, 최종소득세: 20, 최종지방: 21                // AA,AB,AC
  };
  var OUT_WIDTH = 22;

  /* 품의 헤더 이름 → 내부 키 (표기 흔들림 흡수) */
  var HEAD_MAP = [
    [/^일자$/, '일자'],
    [/^소득자\s*명$/, '소득자명'],
    [/^기사\s*코드$/, '기사코드'],
    [/^지급\s*연월$/, '지급연월'],
    [/^소득\s*금액$/, '소득금액'],
    [/^세율$/, '세율'],
    [/^실\s*지급액$/, '실지급액'],
    [/^소득세$/, '소득세'],
    [/^지방\s*소득세$/, '지방소득세'],
    [/^부인\s*신청\s*내역$/, '부인신청내역'],
    [/^변경\s*소득자\s*명$/, '해당자명']
  ];

  function clean(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); }
  function digits(v) { return String(v == null ? '' : v).replace(/[^0-9]/g, ''); }

  /** 헤더 행을 찾아 열 위치를 매핑. 주민등록번호는 순서로 1번째=부인자, 2번째=해당자 */
  function detectLayout(matrix) {
    for (var i = 0; i < Math.min(matrix.length, 12); i++) {
      var row = matrix[i].map(clean);
      if (row.indexOf('소득자명') < 0 && row.indexOf('지급연월') < 0) continue;
      var map = {}, rrnSeen = 0;
      row.forEach(function (h, j) {
        if (/^주민\s*등록\s*번호$/.test(h)) {
          map[rrnSeen === 0 ? '부인주민' : '해당주민'] = j;
          rrnSeen++;
          return;
        }
        for (var k = 0; k < HEAD_MAP.length; k++) {
          if (HEAD_MAP[k][0].test(h)) { map[HEAD_MAP[k][1]] = j; return; }
        }
      });
      if (map['지급연월'] !== undefined && map['소득금액'] !== undefined) {
        return { map: map, headerRow: i, hasHeader: true };
      }
    }
    // 헤더가 없으면 열 개수로 추정 (12열=기사코드 없음 / 13열=있음)
    var w = matrix.length ? Math.max.apply(null, matrix.map(function (r) { return r.length; })) : 0;
    var base = (w >= 13)
      ? { 일자:0,소득자명:1,부인주민:2,기사코드:3,지급연월:4,소득금액:5,세율:6,실지급액:7,소득세:8,지방소득세:9,부인신청내역:10,해당자명:11,해당주민:12 }
      : { 일자:0,소득자명:1,부인주민:2,지급연월:3,소득금액:4,세율:5,실지급액:6,소득세:7,지방소득세:8,부인신청내역:9,해당자명:10,해당주민:11 };
    return { map: base, headerRow: -1, hasHeader: false, guessed: true };
  }

  /* 사전: { 주민번호13 : {name, code} } */
  function buildBook(matrix, book) {
    book = book || {};
    var added = 0;
    matrix.forEach(function (r) {
      // 시트 H~AC 블록 기준 (소득자명/주민번호, 변경소득자명/주민번호)
      [[OUT.소득자명, OUT.부인주민], [OUT.해당자명, OUT.해당주민]].forEach(function (p) {
        var raw = r[p[0]], rrn = digits(r[p[1]]);
        if (!raw || rrn.length !== 13) return;
        var pc = parseNameCode(raw);
        if (!pc.code) return;
        if (!book[rrn]) added++;
        book[rrn] = { name: pc.name, code: pc.code };
      });
    });
    return { book: book, added: added };
  }

  function parseNameCode(v) {
    var s = String(v == null ? '' : v).trim();
    var nums = s.match(/\d+/g);
    var code = nums ? nums[nums.length - 1] : '';
    var name = code ? s.replace(new RegExp('[\\s\\n]*' + code + '[\\s\\n]*$'), '') : s;
    return { name: name.replace(/[\s\n]+/g, ' ').trim(), code: code ? pad6(code) : '' };
  }
  function pad6(c) { c = String(c).trim(); while (/^\d+$/.test(c) && c.length < 6) c = '0' + c; return c; }

  /** 이름+코드를 시트 형식("이름\n코드")으로 */
  function nameWithCode(name, code) {
    name = clean(name);
    return code ? name + '\n' + code : name;
  }

  /* '전액부인' 계열이면 부인자 최종 = 0 */
  function isFullDenial(text) {
    var s = clean(text);
    if (!s) return true;               // 계속 행(빈칸)은 앞 행의 부인내역을 따름
    return /부인/.test(s) && !/부분/.test(s);
  }

  /**
   * 품의 데이터 → 시트 H~AC 붙여넣기 블록
   * @param matrix  품의에서 복사한 2차원 배열
   * @param book    주민번호→소득자코드 사전
   * @param opt     {startRow:number|null}  수식용 시작 행(없으면 값 대신 빈칸)
   */
  function convert(matrix, book, opt) {
    opt = opt || {}; book = book || {};
    var lay = detectLayout(matrix);
    var m = lay.map;
    var startAt = lay.headerRow >= 0 ? lay.headerRow + 1 : 0;

    var out = [], warn = {
      guessedLayout: !!lay.guessed, hasCode: 0, needCode: [],
      groups: 0, rows: 0, noRrn: [], lookup: []
    };
    var curDenial = '', curBuin = null, curHaing = null;

    for (var i = startAt; i < matrix.length; i++) {
      var r = matrix[i];
      if (!r || !r.some(function (c) { return clean(c) !== ''; })) continue;
      function get(key) { var j = m[key]; return j === undefined ? '' : clean(r[j]); }

      var ym = get('지급연월');
      if (!ym) continue;                       // 금액 행이 아님

      // 새 그룹(소득자명이 있는 행)
      if (get('소득자명')) {
        warn.groups++;
        curDenial = get('부인신청내역');
        curBuin  = resolvePerson(get('소득자명'), get('부인주민'), book, warn, '부인자');
        curHaing = resolvePerson(get('해당자명'), get('해당주민'), book, warn, '해당자');
      }
      if (!curBuin) continue;
      warn.rows++;

      var row = new Array(OUT_WIDTH).fill('');
      row[OUT.일자]        = get('일자');
      row[OUT.소득자명]     = get('소득자명') ? nameWithCode(curBuin.name, curBuin.code) : '';
      row[OUT.부인주민]     = get('소득자명') ? get('부인주민') : '';
      row[OUT.기사코드]     = get('기사코드');
      row[OUT.지급연월]     = ym;
      row[OUT.소득금액]     = get('소득금액');
      row[OUT.세율]        = normRate(get('세율'));
      row[OUT.실지급액]     = get('실지급액');
      row[OUT.소득세]       = get('소득세');
      row[OUT.지방소득세]   = get('지방소득세');
      row[OUT.부인신청내역] = get('부인신청내역');

      // 부인자 최종(S,T,U): 전액부인이면 0
      if (isFullDenial(curDenial)) {
        row[OUT.부인최종소득] = '0';
        row[OUT.부인최종소득세] = '0';
        row[OUT.부인최종지방] = '0';
      }

      if (get('소득자명')) {
        row[OUT.해당자명] = nameWithCode(curHaing.name, curHaing.code);
        row[OUT.해당주민] = get('해당주민');
      }

      // 최종(AA,AB,AC): 시작 행을 알면 수식으로, 모르면 빈칸
      if (opt.startRow) {
        var n = opt.startRow + out.length;
        row[OUT.최종소득]   = '=IF(X' + n + '="","",X' + n + '+M' + n + ')';
        row[OUT.최종소득세] = '=IF(Y' + n + '="","",Y' + n + '+P' + n + ')';
        row[OUT.최종지방]   = '=IF(Z' + n + '="","",Z' + n + '+Q' + n + ')';
      }

      out.push(row);

      // ERP에서 조회할 목록(해당자 기존금액) — 중복 제거
      var key = (curHaing.code || curHaing.name) + '|' + ym;
      if (warn.lookup.every(function (x) { return x.key !== key; })) {
        warn.lookup.push({ key: key, name: curHaing.name, code: curHaing.code, ym: ym });
      }
    }
    return { rows: out, warnings: warn, layout: lay };
  }

  function resolvePerson(rawName, rawRrn, book, warn, role) {
    var pc = parseNameCode(rawName);
    var rrn = digits(rawRrn);
    var code = pc.code;
    if (!code && rrn.length === 13 && book[rrn]) code = book[rrn].code;
    if (code) warn.hasCode++;
    else {
      var label = (pc.name || '(이름없음)') + ' (' + role + ')';
      if (warn.needCode.indexOf(label) < 0) warn.needCode.push(label);
    }
    if (rawName && rrn.length !== 13 && warn.noRrn.indexOf(pc.name) < 0) warn.noRrn.push(pc.name);
    return { name: pc.name, code: code, rrn: rrn };
  }

  function normRate(v) {
    var s = clean(v);
    if (!s) return '';
    if (/%/.test(s)) return s;
    var n = Number(s);
    if (!isNaN(n) && n > 0 && n < 1) return Math.round(n * 100) + '%';
    if (!isNaN(n) && n >= 1) return Math.round(n) + '%';
    return s;
  }

  /** 결과를 붙여넣기용 TSV로 (셀 안의 줄바꿈은 따옴표로 감쌈) */
  function toTsv(rows) {
    return rows.map(function (r) {
      return r.map(function (c) {
        var s = String(c == null ? '' : c);
        return /[\t\n"]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
      }).join('\t');
    }).join('\n');
  }

  var API = { convert: convert, buildBook: buildBook, detectLayout: detectLayout,
              toTsv: toTsv, parseNameCode: parseNameCode, OUT: OUT, OUT_WIDTH: OUT_WIDTH };
  if (typeof module !== 'undefined' && module.exports) module.exports = API;
  else root.IntakeCore = API;
})(typeof self !== 'undefined' ? self : this);
