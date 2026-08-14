/* ============================================================================
 *  사업소득 수정신고 관리 앱 — 화면 로직
 *  데이터는 JSON '장부' 파일 하나에 담긴다(로컬 드라이브/공유드라이브).
 *  브라우저 안에서만 처리하며 외부로 전송하지 않는다.
 * ==========================================================================*/
(function () {
  'use strict';
  var $ = function (id) { return document.getElementById(id); };
  var esc = function (s) { var d = document.createElement('div'); d.textContent = s == null ? '' : s; return d.innerHTML; };
  var comma = function (n) {
    if (n == null || n === '') return '';
    return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  };
  var LS_KEY = 'bureung.ledger.v1';

  /* ── 상태 ── */
  var DB = { v: 1, book: {}, cases: [] };
  var fileName = '';          // 현재 장부 파일명
  var fileHandle = null;      // File System Access 핸들(가능한 브라우저)
  var dirty = false;
  var curId = null;
  var filter = 'open';
  var loadedSheets = null, pickedSheet = -1;

  /* ── 단계 정의 (매뉴얼 기준) ── */
  var STEPS = [
    { k: 'erp',      t: 'ERP 수정',           d: '부인자·해당자 소득 정정' },
    { k: 'htSimple', t: '간이지급명세서',      d: '홈택스 · 귀속월별 수정신고' },
    { k: 'htAnnual', t: '지급명세서',          d: '전년도 이전 소득이 있을 때' },
    { k: 'reply',    t: '재무 회신',           d: '기안문 댓글로 완료 회신' }
  ];

  /* ══════════ 저장/불러오기 ══════════ */
  function markDirty(v) {
    dirty = v !== false;
    $('filechip').classList.toggle('dirty', dirty);
    $('filename').textContent = (fileName || '저장 안 됨') + (dirty ? ' • 변경됨' : '');
    if (dirty) backup();
  }
  function backup() {
    try { localStorage.setItem(LS_KEY, JSON.stringify({ fileName: fileName, db: DB })); } catch (e) {}
  }
  function restore() {
    try {
      var raw = localStorage.getItem(LS_KEY);
      if (!raw) return false;
      var o = JSON.parse(raw);
      if (!o || !o.db) return false;
      DB = normalize(o.db); fileName = o.fileName || '';
      return true;
    } catch (e) { return false; }
  }
  function normalize(db) {
    db = db || {};
    db.v = 1; db.book = db.book || {}; db.cases = db.cases || [];
    db.cases.forEach(function (c) {
      c.steps = c.steps || {}; c.groups = c.groups || [];
      c.groups.forEach(function (g) {
        g.months = g.months || [];
        g.months.forEach(function (m) {
          m.bFinal = m.bFinal || { amt: 0, tax: 0, local: 0 };
          m.hPrev = m.hPrev || { amt: null, tax: null, local: null };
        });
      });
    });
    return db;
  }

  async function saveLedger() {
    var text = JSON.stringify(DB, null, 1);
    var name = fileName || ('수정신고장부_' + ymd() + '.json');
    // 같은 파일에 덮어쓰기(지원 브라우저)
    if (fileHandle && fileHandle.createWritable) {
      try {
        var w = await fileHandle.createWritable();
        await w.write(text); await w.close();
        fileName = fileHandle.name; markDirty(false);
        toast('저장했습니다 · ' + fileName);
        return;
      } catch (e) { /* 실패 시 다운로드로 */ }
    }
    if (window.showSaveFilePicker) {
      try {
        var h = await window.showSaveFilePicker({
          suggestedName: name,
          types: [{ description: 'JSON 장부', accept: { 'application/json': ['.json'] } }]
        });
        var w2 = await h.createWritable(); await w2.write(text); await w2.close();
        fileHandle = h; fileName = h.name; markDirty(false);
        toast('저장했습니다 · ' + fileName);
        return;
      } catch (e) {
        if (e && e.name === 'AbortError') return;
      }
    }
    download(new Blob([text], { type: 'application/json' }), name);
    fileName = name; markDirty(false);
    toast('장부 파일을 내려받았습니다 · ' + name);
  }

  async function openLedger() {
    if (window.showOpenFilePicker) {
      try {
        var arr = await window.showOpenFilePicker({
          types: [{ description: 'JSON 장부', accept: { 'application/json': ['.json'] } }]
        });
        fileHandle = arr[0];
        var f = await fileHandle.getFile();
        applyLedger(await f.text(), f.name);
        return;
      } catch (e) {
        if (e && e.name === 'AbortError') return;
      }
    }
    var inp = document.createElement('input');
    inp.type = 'file'; inp.accept = '.json,application/json';
    inp.addEventListener('change', function () {
      var f = inp.files && inp.files[0]; if (!f) return;
      var fr = new FileReader();
      fr.onload = function () { applyLedger(String(fr.result), f.name); };
      fr.readAsText(f);
    });
    inp.click();
  }
  function applyLedger(text, name) {
    try {
      var o = JSON.parse(text);
      if (!o || !Array.isArray(o.cases)) throw new Error('형식이 다릅니다');
      DB = normalize(o); fileName = name;
      markDirty(false); showList();
      toast('장부를 불러왔습니다 · ' + DB.cases.length + '건');
    } catch (e) { alert('장부 파일을 읽지 못했습니다: ' + (e.message || e)); }
  }
  function download(blob, name) {
    // claude.ai 아티팩트 화면은 xlsx·zip 저장을 허용하지 않는다 → 안내로 대체
    if (window.claude && window.claude.downloads) {
      window.claude.downloads.save({ filename: name, data: blob })
        .then(function () { toast('내려받았습니다 · ' + name); })
        .catch(function (e) {
          var c = e && e.code;
          if (c === 'rejected_extension' || c === 'extension_not_enabled') {
            toast('이 링크 화면에서는 파일 저장이 제한됩니다. 내려받은 HTML 파일로 열어 주세요.');
          } else if (c === 'declined') { toast('저장을 취소했습니다.'); }
          else { toast('저장하지 못했습니다: ' + ((e && e.message) || e)); }
        });
      return;
    }
    var url = URL.createObjectURL(blob), a = document.createElement('a');
    a.href = url; a.download = name; document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 1500);
  }
  function toast(msg) {
    var d = document.createElement('div');
    d.textContent = msg;
    d.style.cssText = 'position:fixed;left:50%;bottom:26px;transform:translateX(-50%);background:var(--ink);' +
      'color:var(--bg);padding:10px 18px;border-radius:10px;font-size:13.5px;font-weight:600;z-index:99;' +
      'box-shadow:0 6px 20px rgba(0,0,0,.25);max-width:90vw';
    document.body.appendChild(d);
    setTimeout(function () { d.style.transition = 'opacity .4s'; d.style.opacity = '0'; }, 2200);
    setTimeout(function () { d.remove(); }, 2700);
  }
  function ymd() {
    var d = new Date(), p = function (n) { return ('0' + n).slice(-2); };
    return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate());
  }

  /* ══════════ 계산 ══════════ */
  function finalOf(m) {
    function f(p, b) { return (num(p) == null && num(b) == null) ? null : (num(p) || 0) + (num(b) || 0); }
    return { amt: f(m.hPrev.amt, m.amt), tax: f(m.hPrev.tax, m.tax), local: f(m.hPrev.local, m.local) };
  }
  function num(v) {
    if (v === null || v === undefined || v === '') return null;
    var n = Number(String(v).replace(/[, ]/g, ''));
    return isNaN(n) ? null : n;
  }
  function caseStats(c) {
    var months = 0, people = {}, needAnnual = [], missCode = [], unfilled = 0, taxWarn = [];
    var thisYear = new Date().getFullYear();
    c.groups.forEach(function (g) {
      people[g.buin.rrn || g.buin.name] = 1; people[g.haing.rrn || g.haing.name] = 1;
      if (!g.buin.code) missCode.push(g.buin.name + '(부인자)');
      if (!g.haing.code) missCode.push(g.haing.name + '(해당자)');
      g.months.forEach(function (m) {
        months++;
        var y = parseInt(String(m.ym).slice(0, 4), 10);
        if (y < thisYear && needAnnual.indexOf(y) < 0) needAnnual.push(y);
        if (num(m.hPrev.amt) == null) unfilled++;
        // 2024.01~06 은 소득세·지방세 0원이어야 함
        var mm = parseInt(String(m.ym).slice(4, 6), 10);
        if (y === 2024 && mm >= 1 && mm <= 6 && ((num(m.tax) || 0) > 0 || (num(m.local) || 0) > 0)) {
          taxWarn.push(m.ym);
        }
      });
    });
    needAnnual.sort();
    var done = STEPS.filter(function (s) {
      if (s.k === 'htAnnual' && !needAnnual.length) return true;   // 해당 없음 = 완료로 계산
      return !!c.steps[s.k];
    }).length;
    return { months: months, people: Object.keys(people).length, needAnnual: needAnnual,
      missCode: missCode, unfilled: unfilled, taxWarn: taxWarn,
      pct: Math.round(done / STEPS.length * 100), allDone: done === STEPS.length };
  }

  /* ══════════ 목록 ══════════ */
  function showList() {
    curId = null;
    $('view-list').classList.remove('hidden');
    $('view-case').classList.add('hidden');
    var open = DB.cases.filter(function (c) { return !caseStats(c).allDone; }).length;
    $('list-sub').innerHTML = DB.cases.length
      ? '전체 <b>' + DB.cases.length + '</b>건 · 진행중 <b>' + open + '</b>건'
      : '장부를 불러오거나 아래에서 품의 파일을 올려 시작하세요.';
    var list = DB.cases.filter(function (c) {
      var d = caseStats(c).allDone;
      return filter === 'all' ? true : (filter === 'done' ? d : !d);
    });
    if (!list.length) {
      $('cases').innerHTML = '<div class="empty"><div class="big">🗂️</div>' +
        '<p>' + (DB.cases.length ? '해당하는 건이 없습니다.' : '등록된 수정신고 건이 없습니다.') + '</p></div>';
      return;
    }
    $('cases').innerHTML = list.map(function (c) {
      var s = caseStats(c);
      var chips = STEPS.map(function (st) {
        if (st.k === 'htAnnual' && !s.needAnnual.length) return '<span class="chip na">지급명세서 해당없음</span>';
        return '<span class="chip' + (c.steps[st.k] ? ' ok' : '') + '">' + st.t + (c.steps[st.k] ? ' ✓' : '') + '</span>';
      }).join('');
      var warn = s.unfilled ? '<span class="chip warn">기존금액 ' + s.unfilled + '칸 미입력</span>' : '';
      return '<div class="case' + (s.allDone ? ' done' : '') + '" data-id="' + esc(c.id) + '">' +
        '<div><div class="t">' + esc(c.title) + '</div>' +
        '<div class="m"><span>접수 <b>' + esc(c.receivedAt || '—') + '</b></span>' +
        '<span>소득자 <b>' + s.people + '</b>명</span><span>월 <b>' + s.months + '</b>건</span>' +
        (c.dueAt ? '<span>마감 <b>' + esc(c.dueAt) + '</b></span>' : '') + '</div>' +
        '<div class="chips">' + chips + warn + '</div></div>' +
        '<div class="prog"><div class="bar"><i style="width:' + s.pct + '%"></i></div>' +
        '<span class="pct">' + s.pct + '%</span></div></div>';
    }).join('');
    Array.prototype.forEach.call($('cases').querySelectorAll('.case'), function (el) {
      el.addEventListener('click', function () { showCase(el.getAttribute('data-id')); });
    });
  }

  /* ══════════ 상세 ══════════ */
  function findCase(id) { return DB.cases.filter(function (c) { return c.id === id; })[0]; }

  function showCase(id) {
    var c = findCase(id); if (!c) return showList();
    curId = id;
    $('view-list').classList.add('hidden');
    $('view-case').classList.remove('hidden');
    window.scrollTo({ top: 0, behavior: 'smooth' });
    renderCase();
  }

  function renderCase() {
    var c = findCase(curId); if (!c) return;
    var s = caseStats(c);
    $('c-title').textContent = c.title;
    $('c-sub').innerHTML = '접수 <b>' + esc(c.receivedAt || '—') + '</b> · 소득자 ' + s.people + '명 · 월 ' + s.months + '건'
      + (c.source ? ' · <span style="font-family:var(--mono);font-size:12px">' + esc(c.source) + '</span>' : '')
      + (c.memo ? '<br>' + esc(c.memo) : '');

    var b = '';
    if (s.missCode.length) b += bn('warn', '⚠', '소득자코드 없음: <b>' + esc(s.missCode.join(', ')) + '</b> — ERP <code>소득자등록</code>에서 확인해 아래 이름 옆 코드를 채워야 업로드 파일이 정상 생성됩니다.');
    if (s.unfilled) b += bn('info', '📝', '소득해당자 <b>기존 소득금액 ' + s.unfilled + '칸</b>이 비어 있습니다. ERP에서 조회해 노란 칸에 입력하세요.');
    if (s.needAnnual.length) b += bn('info', '📅', '<b>' + s.needAnnual.join(', ') + '년</b> 소득이 포함되어 있어 간이지급명세서 수정신고 후 <b>지급명세서(연별) 수정신고</b>도 진행해야 합니다.');
    if (s.taxWarn.length) b += bn('danger', '❗', '2024년 1~6월은 소득세·지방소득세가 <b>0원</b>이어야 하는데 값이 있습니다: ' + esc(s.taxWarn.join(', ')) + ' — 품의 내용을 재확인하세요.');
    $('c-banners').innerHTML = b;

    $('c-steps').innerHTML = STEPS.map(function (st) {
      var na = st.k === 'htAnnual' && !s.needAnnual.length;
      var on = na ? true : !!c.steps[st.k];
      return '<button class="step' + (on ? ' on' : '') + (na ? ' na' : '') + '" data-k="' + st.k + '"' + (na ? ' disabled' : '') + '>' +
        '<span class="bx">✓</span><span class="tx"><b>' + st.t + '</b><em>' + (na ? '해당 없음' : st.d) + '</em></span></button>';
    }).join('');
    Array.prototype.forEach.call($('c-steps').querySelectorAll('.step'), function (el) {
      el.addEventListener('click', function () {
        var k = el.getAttribute('data-k');
        c.steps[k] = !c.steps[k]; markDirty(); renderCase();
      });
    });

    renderGroups(c);
    updateGen(c, s);
  }
  function bn(t, ic, msg) { return '<div class="banner ' + t + '"><span class="ic">' + ic + '</span><span>' + msg + '</span></div>'; }

  function renderGroups(c) {
    $('c-groups').innerHTML = c.groups.map(function (g, gi) {
      var rows = g.months.map(function (m, mi) {
        var f = finalOf(m);
        var empty = num(m.hPrev.amt) == null;
        return '<tr>' +
          '<td class="lbl">' + fmtYm(m.ym) + '</td>' +
          '<td>' + comma(m.amt) + '</td>' +
          '<td class="' + ((num(m.bFinal.amt) || 0) === 0 ? 'zero' : '') + ' sep">' + comma(num(m.bFinal.amt) || 0) + '</td>' +
          '<td class="sep">' + inp(gi, mi, 'amt', m.hPrev.amt, empty) + '</td>' +
          '<td>' + inp(gi, mi, 'tax', m.hPrev.tax, empty) + '</td>' +
          '<td>' + inp(gi, mi, 'local', m.hPrev.local, empty) + '</td>' +
          '<td class="calc sep">' + (f.amt == null ? '—' : comma(f.amt)) + '</td>' +
          '<td class="calc">' + (f.tax == null ? '—' : comma(f.tax)) + '</td>' +
          '<td class="calc">' + (f.local == null ? '—' : comma(f.local)) + '</td>' +
          '</tr>';
      }).join('');
      var tot = g.months.reduce(function (a, m) {
        var f = finalOf(m);
        a.d += num(m.amt) || 0; a.f += f.amt || 0; return a;
      }, { d: 0, f: 0 });
      return '<div class="group">' +
        '<div class="ghead">' +
          '<span class="who"><span class="nm">' + esc(g.buin.name) + '</span>' + codeInp(gi, 'buin', g.buin.code) + '</span>' +
          '<span class="arrow">→</span>' +
          '<span class="who"><span class="nm">' + esc(g.haing.name) + '</span>' + codeInp(gi, 'haing', g.haing.code) + '</span>' +
          '<span class="den">' + esc(g.denial || '부인') + '</span>' +
        '</div>' +
        '<div class="tablewrap"><table>' +
        '<thead><tr><th rowspan="2">귀속월</th><th>부인 대상</th><th class="grp">부인자 최종</th>' +
        '<th class="grp" colspan="3">소득해당자 기존 <span style="color:var(--warn)">← ERP 조회</span></th>' +
        '<th class="grp" colspan="3">최종 신고금액 (자동)</th></tr>' +
        '<tr><th>소득금액</th><th class="grp">소득금액</th>' +
        '<th class="grp">소득금액</th><th>소득세</th><th>지방세</th>' +
        '<th class="grp">소득금액</th><th>소득세</th><th>지방세</th></tr></thead>' +
        '<tbody>' + rows + '</tbody>' +
        '<tfoot><tr><td class="lbl">합계</td><td>' + comma(tot.d) + '</td><td class="sep"></td>' +
        '<td class="sep"></td><td></td><td></td><td class="calc sep">' + comma(tot.f) + '</td><td></td><td></td></tr></tfoot>' +
        '</table></div></div>';
    }).join('');

    Array.prototype.forEach.call($('c-groups').querySelectorAll('input.amt'), function (el) {
      el.addEventListener('change', onAmt);
      el.addEventListener('blur', onAmt);
    });
    Array.prototype.forEach.call($('c-groups').querySelectorAll('input.code'), function (el) {
      el.addEventListener('change', onCode);
      el.addEventListener('blur', onCode);
    });
  }
  function codeInp(gi, who, code) {
    return '<input class="code' + (code ? '' : ' miss') + '" data-g="' + gi + '" data-w="' + who + '"' +
      ' value="' + esc(code || '') + '" placeholder="소득자코드" maxlength="6" inputmode="numeric" title="ERP 소득자등록에서 확인한 6자리 코드">';
  }
  function onCode(e) {
    var el = e.target, c = findCase(curId); if (!c) return;
    var g = c.groups[+el.getAttribute('data-g')]; if (!g) return;
    var p = g[el.getAttribute('data-w')];
    var v = String(el.value || '').replace(/[^0-9]/g, '');
    if (v) while (v.length < 6) v = '0' + v;
    if (p.code === v) return;
    p.code = v;
    if (v && p.rrn && p.rrn.length === 13) DB.book[p.rrn] = { name: p.name, code: v };  // 다음에 자동으로
    markDirty(); renderCase();
  }
  function inp(gi, mi, key, val, empty) {
    return '<input class="amt' + (empty ? ' empty' : '') + '" data-g="' + gi + '" data-m="' + mi + '" data-k="' + key + '"' +
      ' inputmode="numeric" value="' + (val == null ? '' : comma(val)) + '" placeholder="0">';
  }
  function codeChip(code) {
    return code ? '<span class="cd">' + esc(code) + '</span>' : '<span class="cd miss">코드없음</span>';
  }
  function fmtYm(ym) {
    ym = String(ym);
    return ym.slice(0, 4) + '.' + ym.slice(4, 6) + '.';
  }
  function onAmt(e) {
    var el = e.target;
    var c = findCase(curId); if (!c) return;
    var g = c.groups[+el.getAttribute('data-g')];
    var m = g && g.months[+el.getAttribute('data-m')];
    if (!m) return;
    var k = el.getAttribute('data-k');
    var v = num(el.value);
    if (m.hPrev[k] === v) return;
    m.hPrev[k] = v;
    // 소득금액만 입력하고 세액을 비워두면 3%로 채워준다(0원 구간 제외)
    if (k === 'amt' && v != null) {
      if (num(m.hPrev.tax) == null && (num(m.tax) || 0) > 0) m.hPrev.tax = Math.floor(v * 0.03);
      if (num(m.hPrev.local) == null && (num(m.local) || 0) > 0) m.hPrev.local = Math.floor(v * 0.003);
    }
    markDirty(); renderCase();
  }

  /* ══════════ 파일 생성 ══════════ */
  function caseToMatrix(c) {
    var out = [];
    c.groups.forEach(function (g) {
      g.months.forEach(function (m, i) {
        var r = new Array(22).fill('');
        var f = finalOf(m);
        if (i === 0) {
          r[0] = c.receivedAt || '';
          r[1] = g.buin.name + (g.buin.code ? '\n' + g.buin.code : '');
          r[2] = g.buin.rrn || '';
          r[3] = g.driver || '';
          r[10] = g.denial || '';
          r[14] = g.haing.name + (g.haing.code ? '\n' + g.haing.code : '');
          r[15] = g.haing.rrn || '';
        }
        r[4] = fmtYm(m.ym);
        r[5] = m.amt == null ? '' : m.amt;
        r[6] = '3%';
        r[8] = m.tax == null ? '' : m.tax;
        r[9] = m.local == null ? '' : m.local;
        r[11] = num(m.bFinal.amt) || 0;
        r[12] = num(m.bFinal.tax) || 0;
        r[13] = num(m.bFinal.local) || 0;
        r[16] = m.hPrev.amt == null ? '' : m.hPrev.amt;
        r[17] = m.hPrev.tax == null ? '' : m.hPrev.tax;
        r[18] = m.hPrev.local == null ? '' : m.hPrev.local;
        r[19] = f.amt == null ? '' : f.amt;
        r[20] = f.tax == null ? '' : f.tax;
        r[21] = f.local == null ? '' : f.local;
        out.push(r);
      });
    });
    return out;
  }
  function tsv(mat) {
    return mat.map(function (r) {
      return r.map(function (c) {
        var s = String(c == null ? '' : c);
        return /[\t\n"]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
      }).join('\t');
    }).join('\n');
  }
  function buildForCase(c) {
    var built = window.ERPCore.buildRows(window.ERPCore.parseClipboard(tsv(caseToMatrix(c))), {});
    var ht = window.HometaxCore.buildHometaxFiles(built, {});
    return { built: built, ht: ht };
  }
  function updateGen(c, s) {
    var r = buildForCase(c);
    $('g-summary').textContent = 'ERP ' + r.built.rows.length + '행 · 홈택스 ' + r.ht.files.length + '개월';
    var notes = [];
    if (s.unfilled) notes.push('⚠️ 기존금액이 비어 있는 칸이 있어 그대로 만들면 금액이 실제와 다를 수 있습니다.');
    if (s.missCode.length) notes.push('⚠️ 소득자코드가 없으면 ERP 업로드가 매칭되지 않습니다.');
    if (r.ht.warnings.noRrn.length) notes.push('⚠️ 주민번호 없는 대상: ' + r.ht.warnings.noRrn.join(', '));
    $('g-note').textContent = notes.join('  ');
  }
  function fileStamp() {
    var d = new Date(), p = function (n) { return ('0' + n).slice(-2); };
    return ymd() + '_' + p(d.getHours()) + p(d.getMinutes());
  }
  function genErp() {
    var c = findCase(curId); var r = buildForCase(c);
    if (!r.built.rows.length) return alert('생성할 데이터가 없습니다.');
    download(new Blob([window.ERPCore.buildXlsx(r.built)],
      { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
      'ERP업로드_' + fileStamp() + '.xlsx');
  }
  function genHt() {
    var c = findCase(curId); var r = buildForCase(c);
    if (!r.ht.files.length) return alert('생성할 데이터가 없습니다.');
    download(new Blob([window.HometaxCore.buildZipOfFiles(r.ht.files, {})], { type: 'application/zip' }),
      '간이지급명세서_' + fileStamp() + '.zip');
  }
  function genAll() {
    var c = findCase(curId); var r = buildForCase(c);
    if (!r.built.rows.length) return alert('생성할 데이터가 없습니다.');
    var entries = [{ name: 'ERP/ERP업로드_' + fileStamp() + '.xlsx', data: window.ERPCore.buildXlsx(r.built) }];
    r.ht.files.forEach(function (f) {
      entries.push({ name: '홈택스/' + f.fileName, data: window.HometaxCore.buildHometaxXlsx(f, {}) });
    });
    download(new Blob([window.HometaxCore.zip(entries)], { type: 'application/zip' }),
      '수정신고_' + safeName(c.title) + '_' + fileStamp() + '.zip');
  }
  function safeName(s) { return String(s || '건').replace(/[\\/:*?"<>|]/g, '').slice(0, 20); }

  /* ══════════ 품의 파일 → 새 건 ══════════ */
  function countRows(mat) {
    return mat.filter(function (r) { return r.some(function (c) { return String(c || '').trim() !== ''; }); }).length;
  }
  async function handleFile(file) {
    if (!file) return;
    var d = $('drop');
    d.querySelector('.t').textContent = '읽는 중… ' + file.name;
    try {
      var sheets = (await window.XlsxReader.read(await file.arrayBuffer()))
        .filter(function (s) { return countRows(s.rows) > 1; });
      if (!sheets.length) throw new Error('내용이 있는 시트를 찾지 못했습니다.');
      loadedSheets = sheets; loadedSheets.fileName = file.name;
      d.querySelector('.ic').textContent = '✅';
      d.querySelector('.t').textContent = file.name;
      d.querySelector('.s').textContent = '시트 ' + sheets.length + '개 · 등록할 품의를 고르세요';
      paintSheets();
    } catch (err) {
      d.querySelector('.ic').textContent = '⚠️';
      d.querySelector('.t').textContent = '파일을 읽지 못했습니다';
      d.querySelector('.s').textContent = String(err && err.message || err);
      loadedSheets = null; $('sheets').classList.add('hidden');
    }
  }
  function paintSheets() {
    if (!loadedSheets) return $('sheets').classList.add('hidden');
    $('sheets').classList.remove('hidden');
    $('sheets').innerHTML = '<span class="lbl">등록할 품의(시트)</span>' +
      loadedSheets.map(function (s, i) {
        return '<button data-i="' + i + '">' + esc(s.name) + ' · ' + countRows(s.rows) + '행</button>';
      }).join('');
    Array.prototype.forEach.call($('sheets').querySelectorAll('button'), function (b) {
      b.addEventListener('click', function () { addCaseFromSheet(+b.getAttribute('data-i')); });
    });
  }
  function addCaseFromSheet(i) {
    var sh = loadedSheets && loadedSheets[i]; if (!sh) return;
    var res = window.IntakeCore.convert(sh.rows, DB.book, {});
    if (!res.rows.length) return alert('이 시트에서 소득 내역을 찾지 못했습니다.');
    var O = window.IntakeCore.OUT;
    var groups = [], cur = null;
    res.rows.forEach(function (r) {
      if (r[O.소득자명]) {
        var pb = window.IntakeCore.parseNameCode(r[O.소득자명]);
        var ph = window.IntakeCore.parseNameCode(r[O.해당자명]);
        cur = {
          buin: { name: pb.name, code: pb.code, rrn: digits(r[O.부인주민]) },
          haing: { name: ph.name, code: ph.code, rrn: digits(r[O.해당주민]) },
          denial: r[O.부인신청내역] || '', driver: r[O.기사코드] || '', months: []
        };
        groups.push(cur);
      }
      if (!cur) return;
      cur.months.push({
        ym: toYm(r[O.지급연월]),
        amt: num(r[O.소득금액]), tax: num(r[O.소득세]), local: num(r[O.지방소득세]),
        bFinal: { amt: num(r[O.부인최종소득]) || 0, tax: num(r[O.부인최종소득세]) || 0, local: num(r[O.부인최종지방]) || 0 },
        hPrev: { amt: null, tax: null, local: null }
      });
    });
    // 사전 보강 (이번 품의에서 확보된 코드 기억)
    groups.forEach(function (g) {
      [g.buin, g.haing].forEach(function (p) {
        if (p.rrn && p.rrn.length === 13 && p.code) DB.book[p.rrn] = { name: p.name, code: p.code };
      });
    });
    var id = 'c' + Date.now().toString(36);
    DB.cases.unshift({
      id: id, title: sh.name + ' 품의', receivedAt: guessDate(res.rows, O), dueAt: '',
      source: (loadedSheets.fileName || '') + ' / ' + sh.name, memo: '',
      groups: groups, steps: {}
    });
    markDirty(); showCase(id);
    toast('새 건을 등록했습니다 · 소득자 ' + groups.length + '건');
  }
  function digits(v) { return String(v == null ? '' : v).replace(/[^0-9]/g, ''); }
  function toYm(v) {
    var m = String(v == null ? '' : v).match(/(\d{4})\D*(\d{1,2})/);
    return m ? m[1] + ('0' + m[2]).slice(-2) : '';
  }
  function guessDate(rows, O) {
    for (var i = 0; i < rows.length; i++) {
      var s = String(rows[i][O.일자] || '');
      var m = s.match(/(\d{2,4})[.\-\/](\d{1,2})[.\-\/](\d{1,2})/);
      if (m) {
        var y = m[1].length === 2 ? '20' + m[1] : m[1];
        return y + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
      }
    }
    return '';
  }

  /* ══════════ 이벤트 ══════════ */
  $('btn-save').addEventListener('click', saveLedger);
  $('btn-open').addEventListener('click', openLedger);
  $('go-home').addEventListener('click', showList);
  $('back').addEventListener('click', showList);
  Array.prototype.forEach.call(document.querySelectorAll('.filters button'), function (b) {
    b.addEventListener('click', function () {
      filter = b.getAttribute('data-f');
      Array.prototype.forEach.call(document.querySelectorAll('.filters button'), function (x) {
        x.classList.toggle('on', x === b);
      });
      showList();
    });
  });
  $('drop').addEventListener('click', function () { $('file').click(); });
  $('drop').addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('file').click(); } });
  $('file').addEventListener('change', function () { handleFile($('file').files && $('file').files[0]); });
  ['dragenter', 'dragover'].forEach(function (ev) {
    $('drop').addEventListener(ev, function (e) { e.preventDefault(); $('drop').classList.add('over'); });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    $('drop').addEventListener(ev, function (e) { e.preventDefault(); $('drop').classList.remove('over'); });
  });
  $('drop').addEventListener('drop', function (e) {
    var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) handleFile(f);
  });
  $('c-edit').addEventListener('click', function () {
    var c = findCase(curId); if (!c) return;
    var box = $('c-editbox');
    box.classList.toggle('hidden');
    if (!box.classList.contains('hidden')) {
      $('e-title').value = c.title; $('e-recv').value = c.receivedAt || '';
      $('e-due').value = c.dueAt || ''; $('e-memo').value = c.memo || '';
    }
  });
  $('e-save').addEventListener('click', function () {
    var c = findCase(curId); if (!c) return;
    c.title = $('e-title').value.trim() || c.title;
    c.receivedAt = $('e-recv').value; c.dueAt = $('e-due').value; c.memo = $('e-memo').value.trim();
    $('c-editbox').classList.add('hidden');
    markDirty(); renderCase();
  });
  $('c-del').addEventListener('click', function () {
    var c = findCase(curId); if (!c) return;
    if (!confirm('“' + c.title + '” 건을 삭제합니다. 되돌릴 수 없습니다.')) return;
    DB.cases = DB.cases.filter(function (x) { return x.id !== curId; });
    markDirty(); showList();
  });
  /* ── 소득자코드 사전 ── */
  function paintBook() {
    var n = Object.keys(DB.book || {}).length;
    $('book-stat').textContent = '등록된 소득자 ' + n + '명';
  }
  $('book-toggle').addEventListener('click', function () { $('book-panel').classList.toggle('hidden'); });
  $('book-add').addEventListener('click', function () {
    var t = $('book-input').value;
    if (!t.trim()) return $('book-input').focus();
    var r = window.IntakeCore.buildBook(window.ERPCore.parseClipboard(t), DB.book);
    DB.book = r.book; $('book-input').value = '';
    markDirty(); paintBook();
    toast('사전에 ' + r.added + '명 추가 · 총 ' + Object.keys(DB.book).length + '명');
  });
  $('book-export').addEventListener('click', function () {
    download(new Blob([JSON.stringify(DB.book, null, 1)], { type: 'application/json' }), '소득자코드사전.json');
  });
  $('book-import').addEventListener('click', function () {
    var inp = document.createElement('input');
    inp.type = 'file'; inp.accept = '.json,application/json';
    inp.addEventListener('change', function () {
      var f = inp.files && inp.files[0]; if (!f) return;
      var fr = new FileReader();
      fr.onload = function () {
        try {
          var o = JSON.parse(String(fr.result)), n = 0;
          Object.keys(o).forEach(function (k) { if (!DB.book[k]) n++; DB.book[k] = o[k]; });
          markDirty(); paintBook(); toast('사전 불러오기 완료 · 새로 ' + n + '명');
        } catch (e) { alert('사전 파일을 읽지 못했습니다.'); }
      };
      fr.readAsText(f);
    });
    inp.click();
  });
  $('book-clear').addEventListener('click', function () {
    if (!confirm('소득자코드 사전을 비웁니다. 계속할까요?')) return;
    DB.book = {}; markDirty(); paintBook();
  });

  $('g-erp').addEventListener('click', genErp);
  $('g-ht').addEventListener('click', genHt);
  $('g-all').addEventListener('click', genAll);
  window.addEventListener('beforeunload', function (e) {
    if (dirty) { e.preventDefault(); e.returnValue = ''; }
  });

  /* ══════════ 시작 ══════════ */
  if (restore()) { markDirty(false); $('filename').textContent = (fileName || '임시 저장본') + ' (브라우저 보관)'; }
  paintBook();
  showList();
  try{ if(window.self!==window.top) $('envwarn').hidden=false; }
  catch(e){ $('envwarn').hidden=false; }
})();
