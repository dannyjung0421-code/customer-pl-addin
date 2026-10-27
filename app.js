/* 2026 거래처별 손익등급 - Excel 추가 기능 (ES5 / IE11 호환)
 * 열려 있는 통합 문서의 시트를 읽어 화면을 그린다.
 *   거래처데이터 : 기준(당월/누적)·월·사업부·영업부·영업부서·고객코드·고객사명·매출·매출총이익·변동비·직접비·판관비·공헌이익·직접이익·영업이익
 *   기준선 P2    : A등급 목표 마진 (없으면 5%p)
 *   제외 영업부서 : 매핑 밖 영업부서 손익 (선택)
 */
(function () {
  'use strict';

  var ORDER = [
    ['TC사업부', [['FC영업부', ['FC1팀', 'FC2팀']], ['CC영업부', ['뉴채널팀', 'CNT플랫폼팀', 'CNT영업팀', '머신운영팀']]]],
    ['소매사업부', [['소매사업부', ['KAM1팀', 'KAM2팀']]]],
    ['소재사업부', [['소재사업부', ['소재영업1팀', '소재영업2팀', '소재사업부(영업)']]]],
    ['생두사업부', [['생두사업부', ['생두운영팀']]]]
  ];
  var GRADES = ['A', 'B', 'C', 'D'];
  var GNAME = { 'A': '목표달성', 'B': '영업이익 BEP', 'C': '직접영업이익 BEP', 'D': '직접영업이익 적자', '등급외': '매출 없음' };
  var GCOL = { 'A': '#1f7a3d', 'B': '#3b6fd8', 'C': '#a86400', 'D': '#c0352b', '등급외': '#9aa1ad' };
  var F = ['매출', '매출총이익', '변동비', '직접비', '판관비', '공헌이익', '직접이익', '영업이익'];
  var TEXT_COLS = { '고객사명': 1, '영업부서': 1, '등급': 1 };

  var S = {
    tab: 'team', mode: 'C', m: 9, maxM: 9, lv: '영업부서', team: 'FC1팀',
    gf: '전체', q: '', sort: '매출', dir: -1, top: true, sorted: false, hlSort: null, hlRows: {}, hlCols: {},
    targetMargin: 0.05, ready: false
  };
  var D = { bep: {}, grades: [], byPeriod: {}, rollup: {}, excluded: {}, view: [] };

  /* ---------------- 공통 ---------------- */
  function $(id) { return document.getElementById(id); }
  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== null && text !== undefined) e.appendChild(document.createTextNode(String(text)));
    return e;
  }
  function clear(n) { while (n.firstChild) n.removeChild(n.firstChild); }
  function num(v) { var n = Number(v); return isFinite(n) ? n : 0; }
  function up(node, attr, stop) {
    while (node && node !== stop && node !== document) {
      if (node.getAttribute && node.getAttribute(attr) !== null) return node;
      node = node.parentNode;
    }
    return null;
  }
  function fixed(v, d) {
    var s = Math.abs(v).toFixed(d);
    var neg = v < 0 && parseFloat(s) !== 0;
    var p = s.split('.');
    p[0] = p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return (neg ? '-' : '') + p.join('.');
  }
  function fInt(v) { return fixed(v, 0); }
  function fMil(v) { return v === null || v === undefined ? '—' : fixed(v / 1e6, 0); }
  function fEok(v) { return v === null || v === undefined ? '—' : fixed(v / 1e8, 1); }
  function fPct(v) { return v === null || v === undefined || !isFinite(v) ? '—' : fixed(v * 100, 1) + '%'; }
  function fPct0(v) { return v === null || v === undefined || !isFinite(v) ? '0%' : fixed(v * 100, 0) + '%'; }
  function negc(v) { return v !== null && v !== undefined && v < 0 ? 'neg' : ''; }
  function td(tr, cls, text) { var c = el('td', cls, text); tr.appendChild(c); return c; }
  function periodLabel() { return S.mode === 'M' ? '2026년 ' + S.m + '월 당월' : '2026년 1~' + S.m + '월 누적'; }
  function showMsg(text, ok) {
    var m = $('msg');
    m.className = 'msg' + (ok ? ' ok' : '');
    m.style.display = text ? 'block' : 'none';
    clear(m); if (text) m.appendChild(document.createTextNode(text));
  }
  function pathOf(team) {
    for (var i = 0; i < ORDER.length; i++) for (var j = 0; j < ORDER[i][1].length; j++)
      if (ORDER[i][1][j][1].indexOf(team) >= 0) return ORDER[i][0] === ORDER[i][1][j][0] ? ORDER[i][0] : ORDER[i][0] + ' > ' + ORDER[i][1][j][0];
    return '';
  }

  /* ---------------- 데이터 읽기 ---------------- */
  function tableFrom(values, firstHeader) {
    var h = -1, i, j;
    for (i = 0; i < Math.min(values.length, 10); i++) if (String(values[i][0]) === firstHeader) { h = i; break; }
    if (h < 0) return [];
    var head = values[h], out = [];
    for (i = h + 1; i < values.length; i++) {
      var r = values[i];
      if (r[0] === '' || r[0] === null) continue;
      var o = {};
      for (j = 0; j < head.length; j++) o[String(head[j])] = r[j];
      out.push(o);
    }
    return out;
  }

  function load() {
    showMsg('');
    $('take').innerHTML = '';
    $('take').appendChild(document.createTextNode('엑셀에서 데이터를 읽는 중입니다…'));
    Excel.run(function (ctx) {
      var sheets = ctx.workbook.worksheets;
      sheets.load('items/name');
      return ctx.sync().then(function () {
        var names = {}, k;
        for (k = 0; k < sheets.items.length; k++) names[sheets.items[k].name] = 1;
        if (!names['거래처데이터']) throw new Error('이 파일에 "거래처데이터" 시트가 없습니다.');
        var r1 = sheets.getItem('거래처데이터').getUsedRange(); r1.load('values');
        var r2 = names['제외 영업부서'] ? sheets.getItem('제외 영업부서').getUsedRange() : null;
        var r3 = names['기준선'] ? sheets.getItem('기준선').getRange('P2') : null;
        if (r2) r2.load('values');
        if (r3) r3.load('values');
        return ctx.sync().then(function () {
          build(r1.values, r2 ? r2.values : [], r3 ? r3.values[0][0] : null);
        });
      });
    }).catch(function (err) {
      $('take').innerHTML = '';
      showMsg('데이터를 읽지 못했습니다. ' + (err && err.message ? err.message : String(err)));
    });
  }

  function build(values, exValues, tm) {
    var rows = tableFrom(values, '기준'), pl = [], ex = [], i, r, o, j;
    for (i = 0; i < rows.length; i++) {
      r = rows[i];
      if (String(r['기준']) !== '당월') continue;
      o = { 월: num(r['월']), 사업부: String(r['사업부']), 영업부: String(r['영업부']), 영업부서: String(r['영업부서']), 고객코드: String(r['고객코드']), 고객사명: String(r['고객사명']) };
      for (j = 0; j < F.length; j++) o[F[j]] = num(r[F[j]]);
      pl.push(o);
    }
    var exRows = tableFrom(exValues, '기준');
    for (i = 0; i < exRows.length; i++) {
      r = exRows[i];
      if (String(r['기준']) !== '당월') continue;
      o = { 월: num(r['월']), 영업부서: String(r['영업부서']), 고객코드: String(r['고객코드']) };
      for (j = 0; j < F.length; j++) o[F[j]] = num(r[F[j]]);
      ex.push(o);
    }
    if (!pl.length) { showMsg('거래처데이터 시트에서 "당월" 행을 찾지 못했습니다.'); return; }
    S.targetMargin = (tm !== null && tm !== '' && isFinite(Number(tm))) ? Number(tm) : 0.05;
    var maxM = 1;
    for (i = 0; i < pl.length; i++) if (pl[i].월 > maxM) maxM = pl[i].월;
    S.maxM = maxM;
    if (!S.ready || S.m > maxM) S.m = maxM;
    compute(pl, ex);
    S.ready = true;
    fillMonths();
    render();
  }

  /* ---------------- 계산 ---------------- */
  function compute(pl, ex) {
    var i, k, r, key, a, maxM = S.maxM;
    // 기준선
    var bep = {};
    for (i = 0; i < pl.length; i++) {
      r = pl[i];
      for (k = r.월; k <= maxM; k++) {
        addBep(bep, 'C|' + k + '|' + r.영업부서, r);
      }
      addBep(bep, 'M|' + r.월 + '|' + r.영업부서, r);
    }
    for (key in bep) if (bep.hasOwnProperty(key)) {
      a = bep[key];
      a.영업BEP = a.매출 ? (a.판관비 - a.변동비) / a.매출 : null;
      a.직접BEP = a.매출 ? (a.직접비 - a.변동비) / a.매출 : null;
      a.목표 = a.영업BEP === null ? null : a.영업BEP + S.targetMargin;
      a.공헌이익률 = a.매출 ? a.공헌이익 / a.매출 : null;
    }
    D.bep = bep;

    // 거래처 등급
    var cust = {}, ids = [];
    for (i = 0; i < pl.length; i++) {
      r = pl[i];
      key = r.영업부서 + '|' + r.고객코드;
      var c = cust[key];
      if (!c) { c = { 사업부: r.사업부, 영업부: r.영업부, 영업부서: r.영업부서, 고객코드: r.고객코드, 고객사명: r.고객사명, mon: {} }; cust[key] = c; ids.push(key); }
      var v = c.mon[r.월];
      if (!v) { v = [0, 0, 0, 0, 0, 0, 0, 0]; c.mon[r.월] = v; }
      for (k = 0; k < F.length; k++) v[k] += r[F[k]];
    }
    var grades = [], byPeriod = {};
    for (i = 0; i < ids.length; i++) {
      var cu = cust[ids[i]], mg = [], cum = [0, 0, 0, 0, 0, 0, 0, 0], seen = false, m, j;
      for (m = 1; m <= 9; m++) mg.push(cu.mon[m] ? gradeOf(cu.mon[m], bep['M|' + m + '|' + cu.영업부서]) : '');
      for (m = 1; m <= maxM; m++) {
        if (cu.mon[m]) {
          for (j = 0; j < 8; j++) cum[j] += cu.mon[m][j];
          seen = true;
          pushRow(grades, byPeriod, cu, 'M', m, cu.mon[m], mg[m - 1], mg);
        }
        if (seen) pushRow(grades, byPeriod, cu, 'C', m, cum.slice(0), gradeOf(cum, bep['C|' + m + '|' + cu.영업부서]), mg);
      }
    }
    D.grades = grades; D.byPeriod = byPeriod;

    // 조직·등급 집계
    var roll = {};
    for (i = 0; i < grades.length; i++) {
      r = grades[i];
      var units = [['전사', '전사'], ['사업부', r.사업부], ['영업부', r.영업부], ['영업부서', r.영업부서]];
      for (k = 0; k < units.length; k++) {
        addRoll(roll, r.기준 + '|' + r.월 + '|' + units[k][0] + '|' + units[k][1] + '|' + r.등급, r);
        addRoll(roll, r.기준 + '|' + r.월 + '|' + units[k][0] + '|' + units[k][1] + '|전체', r);
      }
    }
    for (key in roll) if (roll.hasOwnProperty(key)) {
      a = roll[key];
      var tot = roll[key.substring(0, key.lastIndexOf('|')) + '|전체'];
      a.매출총이익률 = a.매출 ? a.매출총이익 / a.매출 : null;
      a.공헌이익률 = a.매출 ? a.공헌이익 / a.매출 : null;
      a.직접이익률 = a.매출 ? a.직접이익 / a.매출 : null;
      a.영업이익률 = a.매출 ? a.영업이익 / a.매출 : null;
      a.매출비중 = tot.매출 ? a.매출 / tot.매출 : null;
      a.거래처비중 = tot.거래처수 ? a.거래처수 / tot.거래처수 : null;
    }
    D.rollup = roll;

    // 제외 영업부서
    var exa = {};
    for (i = 0; i < ex.length; i++) {
      r = ex[i];
      addEx(exa, 'M|' + r.월, r);
      for (k = r.월; k <= maxM; k++) addEx(exa, 'C|' + k, r);
    }
    D.excluded = exa;
  }
  function addBep(bep, key, r) {
    var a = bep[key];
    if (!a) { a = { 사업부: r.사업부, 영업부서: r.영업부서, 매출: 0, 변동비: 0, 직접비: 0, 판관비: 0, 공헌이익: 0 }; bep[key] = a; }
    a.매출 += r.매출; a.변동비 += r.변동비; a.직접비 += r.직접비; a.판관비 += r.판관비; a.공헌이익 += r.공헌이익;
  }
  function gradeOf(v, t) {
    if (!t || v[0] <= 0) return '등급외';
    var rate = v[5] / v[0];
    return rate >= t.목표 ? 'A' : rate >= t.영업BEP ? 'B' : rate >= t.직접BEP ? 'C' : 'D';
  }
  function pushRow(arr, idx, c, mode, m, v, g, mg) {
    var o = {
      key: mode + '|' + m + '|' + c.영업부서 + '|' + c.고객코드, 기준: mode, 월: m, 사업부: c.사업부, 영업부: c.영업부, 영업부서: c.영업부서,
      고객코드: c.고객코드, 고객사명: c.고객사명, 매출: v[0], 매출총이익: v[1], 공헌이익: v[5], 직접이익: v[6], 영업이익: v[7],
      매출총이익률: v[0] > 0 ? v[1] / v[0] : null, 공헌이익률: v[0] > 0 ? v[5] / v[0] : null,
      직접이익률: v[0] > 0 ? v[6] / v[0] : null, 영업이익률: v[0] > 0 ? v[7] / v[0] : null, 등급: g, mg: mg
    };
    arr.push(o);
    var p = mode + '|' + m;
    if (!idx[p]) idx[p] = [];
    idx[p].push(o);
  }
  function addRoll(roll, key, r) {
    var a = roll[key];
    if (!a) { a = { 거래처수: 0, 매출: 0, 매출총이익: 0, 공헌이익: 0, 직접이익: 0, 영업이익: 0 }; roll[key] = a; }
    a.거래처수 += 1; a.매출 += r.매출; a.매출총이익 += r.매출총이익; a.공헌이익 += r.공헌이익; a.직접이익 += r.직접이익; a.영업이익 += r.영업이익;
  }
  function addEx(exa, p, r) {
    if (!exa[p]) exa[p] = {};
    var a = exa[p][r.영업부서];
    if (!a) { a = { 영업부서: r.영업부서, codes: {}, 거래처수: 0, 매출: 0, 매출총이익: 0, 공헌이익: 0, 직접이익: 0, 영업이익: 0 }; exa[p][r.영업부서] = a; }
    if (!a.codes[r.고객코드]) { a.codes[r.고객코드] = 1; a.거래처수 += 1; }
    a.매출 += r.매출; a.매출총이익 += r.매출총이익; a.공헌이익 += r.공헌이익; a.직접이익 += r.직접이익; a.영업이익 += r.영업이익;
  }
  function R(unit, name, g) { return D.rollup[S.mode + '|' + S.m + '|' + unit + '|' + name + '|' + (g || '전체')]; }

  /* ---------------- 공통 화면 ---------------- */
  function fillMonths() {
    var sel = $('monthSel'); clear(sel);
    for (var m = 1; m <= S.maxM; m++) { var o = el('option', null, m + '월'); o.value = String(m); sel.appendChild(o); }
    sel.value = String(S.m);
  }
  function setOn(container, attr, value) {
    var bs = container.getElementsByTagName('button');
    for (var i = 0; i < bs.length; i++) {
      var on = bs[i].getAttribute(attr) === String(value);
      bs[i].className = bs[i].className.replace(/\s*\bon\b/g, '') + (on ? ' on' : '');
    }
  }
  function renderTake() {
    var t = $('take'); clear(t);
    var all = R('전사', '전사');
    if (!all) return;
    t.appendChild(el('span', 'tk', periodLabel() + ' · 전체 ' + fInt(all.거래처수) + '곳'));
    for (var i = 0; i < GRADES.length; i++) {
      var g = GRADES[i], r = R('전사', '전사', g);
      t.appendChild(el('span', 'sep', '|'));
      var s = el('span', 'tk');
      var sw = el('span', 'sw'); sw.style.backgroundColor = GCOL[g]; s.appendChild(sw);
      s.appendChild(document.createTextNode(g + '등급: ' + fInt(r ? r.거래처수 : 0) + '곳 (매출비중 ' + fPct(r ? r.매출비중 : 0) + ')'));
      t.appendChild(s);
    }
  }
  function kpis(host, unit, name, mil) {
    clear(host);
    var r = R(unit, name);
    var fm = function (v) { return mil ? fMil(v) + '백만' : fEok(v) + '억'; };
    var items = [
      ['매출', r ? fm(r.매출) : '—', null, r && r.매출],
      ['공헌이익', r ? fm(r.공헌이익) : '—', r ? '공헌이익률 ' + fPct(r.공헌이익률) : '', r && r.공헌이익],
      ['직접영업이익', r ? fm(r.직접이익) : '—', r ? '직접영업이익률 ' + fPct(r.직접이익률) : '', r && r.직접이익],
      ['영업이익', r ? fm(r.영업이익) : '—', r ? '영업이익률 ' + fPct(r.영업이익률) : '', r && r.영업이익],
      ['거래처', r ? fInt(r.거래처수) + '곳' : '—', null, 0]
    ];
    for (var i = 0; i < items.length; i++) {
      var b = el('div', 'kpi');
      b.appendChild(el('div', 'lab', items[i][0]));
      b.appendChild(el('div', 'val ' + negc(items[i][3]), items[i][1]));
      if (items[i][2] !== null) b.appendChild(el('div', 'sub', items[i][2]));
      host.appendChild(b);
    }
  }
  function gradeCards(host, unit, name) {
    clear(host);
    for (var i = 0; i < GRADES.length; i++) {
      var g = GRADES[i], r = R(unit, name, g);
      var c = el('div', 'gcard'); c.style.borderTopColor = GCOL[g];
      var h = el('div', 'gh'); var l = el('span', 'gl', g); l.style.color = GCOL[g]; h.appendChild(l); h.appendChild(el('span', 'gd', GNAME[g])); c.appendChild(h);
      var tb = el('table'), rows = [
        ['거래처', r ? fInt(r.거래처수) + '곳' : '0곳', 'big', 0],
        ['매출', r ? fEok(r.매출) + '억' : '—', '', 0],
        ['매출 비중', r ? fPct(r.매출비중) : '—', '', 0],
        ['공헌이익률', r ? fPct(r.공헌이익률) : '—', '', r && r.공헌이익률],
        ['직접영업이익률', r ? fPct(r.직접이익률) : '—', '', r && r.직접이익률]
      ];
      for (var j = 0; j < rows.length; j++) {
        var tr = el('tr'); td(tr, 'k', rows[j][0]); td(tr, 'v ' + rows[j][2] + ' ' + negc(rows[j][3]), rows[j][1]); tb.appendChild(tr);
      }
      c.appendChild(tb);
      host.appendChild(c);
    }
  }

  /* ---------------- 조직별 거래처 ---------------- */
  function buildSlicer() {
    var host = $('slicer'); clear(host);
    for (var i = 0; i < ORDER.length; i++) {
      var bu = ORDER[i][0], subs = ORDER[i][1], grp = el('div', 'grp');
      grp.appendChild(slBtn('lv1', bu, '사업부'));
      for (var j = 0; j < subs.length; j++) {
        if (subs[j][0] !== bu) grp.appendChild(slBtn('lv2', subs[j][0], '영업부'));
        for (var k = 0; k < subs[j][1].length; k++) grp.appendChild(slBtn('', subs[j][1][k], '영업부서'));
      }
      host.appendChild(grp);
    }
  }
  function slBtn(cls, name, lv) {
    var b = el('button', cls, name); b.type = 'button';
    b.setAttribute('data-team', name); b.setAttribute('data-lv', lv);
    return b;
  }
  var COLS = [
    ['고객사명', '거래처', 'l', 'text'], ['영업부서', '영업부서', 'l', 'text'], ['등급', '등급', '', 'grade'],
    ['매출', '매출', '', 'mil'], ['매출총이익', '매출총이익', '', 'mil'], ['매출총이익률', '매출총이익률', '', 'pct'],
    ['공헌이익', '공헌이익', '', 'mil'], ['공헌이익률', '공헌이익률', '', 'pct'], ['직접이익', '직접이익', '', 'mil'],
    ['직접이익률', '직접이익률', '', 'pct'], ['영업이익', '영업이익', '', 'mil'], ['영업이익률', '영업이익률', '', 'pct'],
    ['월별', '월별 등급 (1→9월)', 'l', 'strip']
  ];
  function renderTeam() {
    var bs = $('slicer').getElementsByTagName('button');
    for (var i = 0; i < bs.length; i++) {
      var on = bs[i].getAttribute('data-team') === S.team && bs[i].getAttribute('data-lv') === S.lv;
      bs[i].className = bs[i].className.replace(/\s*\bon\b/g, '') + (on ? ' on' : '');
    }
    $('teamTitle').innerHTML = ''; $('teamTitle').appendChild(document.createTextNode(S.team + ' 손익 · ' + periodLabel()));
    $('teamPath').innerHTML = ''; $('teamPath').appendChild(document.createTextNode(S.lv === '영업부서' ? pathOf(S.team) : S.lv === '영업부' ? 'TC사업부' : ''));
    kpis($('teamKpis'), S.lv, S.team, true);
    gradeCards($('teamGrades'), S.lv, S.team);
    var gc = $('gradeChips'); clear(gc);
    var gl = [['전체', '전체 등급'], ['AB', 'A+B'], ['CD', 'C+D'], ['A', 'A ' + GNAME.A], ['B', 'B ' + GNAME.B], ['C', 'C ' + GNAME.C], ['D', 'D ' + GNAME.D], ['등급외', '등급외']];
    for (i = 0; i < gl.length; i++) {
      var b = el('button', S.gf === gl[i][0] ? 'on' : '', gl[i][1]); b.type = 'button'; b.setAttribute('data-gf', gl[i][0]); gc.appendChild(b);
    }
    setOn($('topSeg'), 'data-top', S.top ? '1' : '0');
    renderTable();
  }
  function currentRows() {
    var all = D.byPeriod[S.mode + '|' + S.m] || [], rows = [], i, r, fld = S.lv;
    var q = S.q.toLowerCase();
    for (i = 0; i < all.length; i++) {
      r = all[i];
      if (r[fld] !== S.team) continue;
      if (S.gf !== '전체') {
        if (S.gf.length === 2 && S.gf !== '등급') { if (S.gf.indexOf(r.등급) < 0 || r.등급 === '등급외') continue; }
        else if (r.등급 !== S.gf) continue;
      }
      if (q && (r.고객사명 + ' ' + r.고객코드).toLowerCase().indexOf(q) < 0) continue;
      rows.push(r);
    }
    var total = rows.length;
    if (S.top) { rows.sort(function (a, b) { return b.매출 - a.매출; }); rows = rows.slice(0, 10); }
    var k = S.sort, dir = S.dir;
    rows.sort(function (a, b) {
      var va = a[k], vb = b[k];
      if (va === null || va === undefined) return 1;
      if (vb === null || vb === undefined) return -1;
      if (typeof va === 'string') return va.localeCompare(vb) * dir;
      return (va - vb) * dir;
    });
    return { rows: rows, total: total };
  }
  function renderTable() {
    var oneTeam = S.lv === '영업부서', head = $('custHead'), body = $('custBody'), i, j;
    clear(head); clear(body);
    var cols = [];
    for (i = 0; i < COLS.length; i++) if (!(oneTeam && COLS[i][0] === '영업부서')) cols.push(COLS[i]);
    for (i = 0; i < cols.length; i++) {
      var c = cols[i], th = el('th', (c[2] || '') + (c[3] !== 'strip' ? ' sortable' : ''));
      th.setAttribute('data-col', c[0]);
      var arrow = c[0] === S.sort ? (S.dir > 0 ? ' △' : ' ▽') : '';
      th.appendChild(document.createTextNode(c[1] + arrow));
      if (S.hlCols[c[0]] || S.hlSort === c[0]) th.className += ' hl1';
      head.appendChild(th);
    }
    var res = currentRows();
    D.view = res.rows;
    for (i = 0; i < res.rows.length; i++) {
      var r = res.rows[i], tr = el('tr'), rowOn = !!S.hlRows[r.고객코드];
      for (j = 0; j < cols.length; j++) {
        var cc = cols[j], v = r[cc[0]], cell;
        if (cc[3] === 'text') { cell = td(tr, 'l' + (cc[0] === '고객사명' ? ' namecell' : ''), v); if (cc[0] === '고객사명') { cell.setAttribute('data-code', r.고객코드); cell.title = r.고객코드 + ' · 누르면 행 음영'; } }
        else if (cc[3] === 'grade') { cell = td(tr, ''); var p = el('span', 'pill', r.등급 === '등급외' ? '-' : r.등급); p.style.backgroundColor = GCOL[r.등급]; cell.appendChild(p); }
        else if (cc[3] === 'mil') cell = td(tr, negc(v), fMil(v));
        else if (cc[3] === 'pct') cell = td(tr, negc(v), fPct(v));
        else {
          cell = td(tr, 'l strip');
          for (var mm = 0; mm < 9; mm++) {
            var gv = r.mg[mm], s = el('span', gv ? '' : 'none', gv === '등급외' ? '-' : (gv || '.'));
            if (gv) s.style.backgroundColor = GCOL[gv];
            if (mm + 1 > S.m) s.style.opacity = '0.35';
            s.title = (mm + 1) + '월 ' + (gv ? (gv === '등급외' ? '등급외' : gv + ' ' + GNAME[gv]) : '거래 없음');
            cell.appendChild(s);
          }
        }
        var n = (rowOn ? 1 : 0) + ((S.hlCols[cc[0]] || S.hlSort === cc[0]) ? 1 : 0);
        if (n) cell.className += ' hl' + n;
      }
      body.appendChild(tr);
    }
    var note = (S.top ? '매출 상위 ' + fInt(res.rows.length) + '곳 / 전체 ' + fInt(res.total) + '곳' : fInt(res.rows.length) + '곳 표시') +
      ' · 금액 단위: 백만원 · 항목 이름을 누르면 정렬(숫자는 내림차순부터)과 열 음영, 거래처명을 누르면 행 음영 (겹치면 진하게)';
    $('tableNote').innerHTML = ''; $('tableNote').appendChild(document.createTextNode(note));
  }

  /* ---------------- 요약 ---------------- */
  function renderSummary() {
    kpis($('sumKpis'), '전사', '전사', false);
    gradeCards($('sumGrades'), '전사', '전사');
    var lg = $('legend'); clear(lg);
    for (var i = 0; i < GRADES.length; i++) { var s = el('span'); var ic = el('i'); ic.style.backgroundColor = GCOL[GRADES[i]]; s.appendChild(ic); s.appendChild(document.createTextNode(GRADES[i] + ' ' + GNAME[GRADES[i]])); lg.appendChild(s); }
    var items = [];
    for (i = 0; i < ORDER.length; i++) {
      items.push(['사업부', ORDER[i][0], ORDER[i][0], '']);
      if (ORDER[i][1].length > 1) for (var j = 0; j < ORDER[i][1].length; j++) items.push(['영업부', ORDER[i][1][j][0], ORDER[i][1][j][0], 'child']);
    }
    items.push(['전사', '전사', '분석대상 전체', 'total']);
    var host = $('stack'); clear(host);
    for (i = 0; i < items.length; i++) {
      var row = el('div', 'srow'), nm = el('div', 'nm ' + items[i][3], items[i][2]), bar = el('div', 'sbar');
      bar.setAttribute('data-unit', items[i][0]); bar.setAttribute('data-name', items[i][1]);
      for (var g = 0; g < GRADES.length; g++) {
        var r = R(items[i][0], items[i][1], GRADES[g]), sh = r && r.매출비중 > 0 ? r.매출비중 : 0;
        if (!sh) continue;
        var seg = el('span', null, sh >= 0.08 ? GRADES[g] + ' ' + fPct0(sh) : '');
        seg.style.width = (sh * 100).toFixed(2) + '%'; seg.style.backgroundColor = GCOL[GRADES[g]];
        seg.title = items[i][2] + ' ' + GRADES[g] + ' · ' + fInt(r.거래처수) + '곳 · 매출 ' + fEok(r.매출) + '억 (' + fPct(sh) + ')';
        bar.appendChild(seg);
      }
      row.appendChild(nm); row.appendChild(bar); host.appendChild(row);
    }
    // matrix
    var body = $('matrixBody'); clear(body);
    for (i = 0; i < ORDER.length; i++) {
      var bu = ORDER[i][0];
      mxRow(body, '사업부', bu, bu, 0);
      for (j = 0; j < ORDER[i][1].length; j++) {
        var sub = ORDER[i][1][j], nested = sub[0] !== bu;
        if (nested) mxRow(body, '영업부', sub[0], sub[0], 1);
        for (var t = 0; t < sub[1].length; t++) mxRow(body, '영업부서', sub[1][t], sub[1][t], nested ? 2 : 1);
      }
    }
    mxRow(body, '전사', '전사', '분석대상 전체', -1);
  }
  function mxRow(body, unit, name, label, lv) {
    var r = R(unit, name);
    if (!r) return;
    var tr = el('tr', (lv >= 0 ? 'lv' + lv + ' click' : 'tot'));
    if (unit !== '전사') { tr.setAttribute('data-unit', unit); tr.setAttribute('data-name', name); }
    td(tr, 'l', label); td(tr, negc(r.매출), fEok(r.매출)); td(tr, negc(r.매출총이익), fEok(r.매출총이익)); td(tr, negc(r.매출총이익률), fPct(r.매출총이익률));
    td(tr, negc(r.공헌이익률), fPct(r.공헌이익률)); td(tr, negc(r.직접이익률), fPct(r.직접이익률)); td(tr, negc(r.영업이익률), fPct(r.영업이익률)); td(tr, '', fInt(r.거래처수));
    for (var g = 0; g < GRADES.length; g++) { var x = R(unit, name, GRADES[g]); td(tr, '', fInt(x ? x.거래처수 : 0) + ' (' + fPct0(x ? x.매출비중 : 0) + ')'); }
    body.appendChild(tr);
  }

  /* ---------------- 등급 기준 / 제외 ---------------- */
  function renderRule() {
    var defs = $('defs'); clear(defs);
    var list = [
      ['A 목표달성', '공헌이익률 ≥ 영업이익 BEP + ' + fixed(S.targetMargin * 100, 1) + '%p (기준선 시트 P2)', GCOL.A],
      ['B 영업이익 BEP', '공헌이익률 ≥ (판관비 − 변동비) ÷ 매출', GCOL.B],
      ['C 직접영업이익 BEP', '공헌이익률 ≥ (직접비 − 변동비) ÷ 매출', GCOL.C],
      ['D 직접영업이익 적자', '공헌이익률이 직접영업이익 BEP에 못 미침', GCOL.D],
      ['등급외', '해당 기간 매출이 0 이하', null],
      ['직접비', '[판관] 인건비 ~ 기타판관비 합계', null],
      ['간접/배부비', '[판관] OP1 ~ OP3배부비 합계', null],
      ['변동비', '운반비 + 광고선전비 + 판매촉진비 + 판매수수료 (직접·간접 포함)', null],
      ['공헌이익', '매출총이익 − 변동비', null],
      ['직접영업이익', '매출총이익 − 직접비', null],
      ['영업이익', '매출총이익 − 판관비', null]
    ];
    for (var i = 0; i < list.length; i++) {
      var tr = el('tr'); var k = td(tr, 'k', list[i][0]); if (list[i][2]) k.style.color = list[i][2]; td(tr, 'v', list[i][1]); defs.appendChild(tr);
    }
    $('bepTitle').innerHTML = ''; $('bepTitle').appendChild(document.createTextNode('영업부서별 기준선 · ' + periodLabel()));
    var body = $('bepBody'); clear(body);
    for (i = 0; i < ORDER.length; i++) for (var j = 0; j < ORDER[i][1].length; j++) for (var t = 0; t < ORDER[i][1][j][1].length; t++) {
      var team = ORDER[i][1][j][1][t], b = D.bep[S.mode + '|' + S.m + '|' + team];
      if (!b) continue;
      var r = el('tr', 'click'); r.setAttribute('data-unit', '영업부서'); r.setAttribute('data-name', team);
      td(r, 'l', ORDER[i][0]); td(r, 'l', team); td(r, '', fEok(b.매출)); td(r, negc(b.공헌이익률), fPct(b.공헌이익률)); td(r, '', fPct(b.목표)); td(r, '', fPct(b.영업BEP)); td(r, '', fPct(b.직접BEP));
      body.appendChild(r);
    }
  }
  function renderExcl() {
    var body = $('exclBody'); clear(body);
    var p = D.excluded[S.mode + '|' + S.m] || {}, list = [], k;
    for (k in p) if (p.hasOwnProperty(k)) list.push(p[k]);
    list.sort(function (a, b) { return b.매출 - a.매출; });
    if (!list.length) { var tr0 = el('tr'); var c0 = td(tr0, 'l', '이 기간에는 제외 영업부서 실적이 없습니다 (또는 "제외 영업부서" 시트가 없습니다).'); c0.colSpan = 8; body.appendChild(tr0); return; }
    for (var i = 0; i < list.length; i++) {
      var r = list[i], tr = el('tr');
      td(tr, 'l', r.영업부서); td(tr, '', fInt(r.거래처수)); td(tr, negc(r.매출), fEok(r.매출)); td(tr, negc(r.매출총이익), fEok(r.매출총이익));
      td(tr, negc(r.공헌이익), fEok(r.공헌이익)); td(tr, negc(r.직접이익), fEok(r.직접이익)); td(tr, negc(r.영업이익), fEok(r.영업이익));
      td(tr, negc(r.영업이익), fPct(r.매출 ? r.영업이익 / r.매출 : null));
      body.appendChild(tr);
    }
  }

  /* ---------------- 그리기 ---------------- */
  function render() {
    if (!S.ready) return;
    setOn($('tabs'), 'data-tab', S.tab);
    setOn($('modeSeg'), 'data-mode', S.mode);
    $('monthSel').value = String(S.m);
    var pages = ['team', 'summary', 'rule', 'excl'];
    for (var i = 0; i < pages.length; i++) $('pg-' + pages[i]).style.display = pages[i] === S.tab ? 'block' : 'none';
    renderTake();
    if (S.tab === 'team') renderTeam();
    else if (S.tab === 'summary') renderSummary();
    else if (S.tab === 'rule') renderRule();
    else renderExcl();
  }
  function goUnit(unit, name) {
    if (unit === '전사') return;
    S.tab = 'team'; S.lv = unit; S.team = name; render();
  }
  function resetTable() {
    S.gf = '전체'; S.q = ''; S.sort = '매출'; S.dir = -1; S.top = true; S.sorted = false; S.hlSort = null; S.hlRows = {}; S.hlCols = {};
    $('search').value = '';
  }

  /* ---------------- 시트로 내보내기 ---------------- */
  function exportSheet() {
    var rows = D.view || [];
    var oneTeam = S.lv === '영업부서';
    var head = ['순위', '거래처', '고객코드', '영업부서', '등급', '매출', '매출총이익', '매출총이익률', '공헌이익', '공헌이익률', '직접이익', '직접이익률', '영업이익', '영업이익률'];
    var keys = [null, '고객사명', '고객코드', '영업부서', '등급', '매출', '매출총이익', '매출총이익률', '공헌이익', '공헌이익률', '직접이익', '직접이익률', '영업이익', '영업이익률'];
    var fmts = ['0', '@', '@', '@', '@', '#,##0,,', '#,##0,,', '0.0%', '#,##0,,', '0.0%', '#,##0,,', '0.0%', '#,##0,,', '0.0%'];
    var vals = [head], nf = [fmts.slice(0)], i, j;
    for (i = 0; i < rows.length; i++) {
      var line = [];
      for (j = 0; j < keys.length; j++) {
        var v = keys[j] === null ? i + 1 : rows[i][keys[j]];
        line.push(v === null || v === undefined ? '' : v);
      }
      vals.push(line); nf.push(fmts.slice(0));
    }
    var title = S.team + ' 거래처별 손익 · ' + periodLabel() + ' · 등급 ' + (S.gf === 'AB' ? 'A+B' : S.gf === 'CD' ? 'C+D' : S.gf) + (S.top ? ' · 매출 상위 10곳' : '') + (S.q ? ' · 검색 "' + S.q + '"' : '') + ' (금액 단위: 백만원)';
    var name = '손익등급_내보내기';
    var lastCol = String.fromCharCode(64 + head.length);
    Excel.run(function (ctx) {
      var sheets = ctx.workbook.worksheets;
      sheets.load('items/name');
      return ctx.sync().then(function () {
        for (var k = 0; k < sheets.items.length; k++) if (sheets.items[k].name === name) sheets.items[k].delete();
        var ws = sheets.add(name);
        ws.getRange('A1').values = [[title]];
        ws.getRange('A1').format.font.bold = true;
        ws.getRange('A1').format.font.size = 12;
        var rg = ws.getRange('A3:' + lastCol + (3 + vals.length - 1));
        rg.numberFormat = nf;
        rg.values = vals;
        var hd = ws.getRange('A3:' + lastCol + '3');
        hd.format.font.bold = true; hd.format.font.color = '#FFFFFF'; hd.format.fill.color = '#1F3864';
        rg.format.font.name = '맑은 고딕';
        rg.format.autofitColumns();
        ws.activate();
        return ctx.sync();
      });
    }).then(function () {
      showMsg('"' + name + '" 시트에 ' + rows.length + '곳을 내보냈습니다. 저장하면 회사 문서보안이 적용됩니다.', true);
    }).catch(function (err) {
      showMsg('내보내지 못했습니다. ' + (err && err.message ? err.message : String(err)));
    });
  }

  /* ---------------- 이벤트 ---------------- */
  function bind() {
    $('tabs').addEventListener('click', function (e) { var b = up(e.target, 'data-tab', this); if (b) { S.tab = b.getAttribute('data-tab'); render(); } });
    $('modeSeg').addEventListener('click', function (e) { var b = up(e.target, 'data-mode', this); if (b) { S.mode = b.getAttribute('data-mode'); render(); } });
    $('monthSel').addEventListener('change', function () { S.m = Number(this.value); render(); });
    $('reload').addEventListener('click', function () { load(); });
    $('slicer').addEventListener('click', function (e) { var b = up(e.target, 'data-team', this); if (b) { S.lv = b.getAttribute('data-lv'); S.team = b.getAttribute('data-team'); renderTeam(); } });
    $('gradeChips').addEventListener('click', function (e) { var b = up(e.target, 'data-gf', this); if (b) { S.gf = b.getAttribute('data-gf'); renderTeam(); } });
    $('topSeg').addEventListener('click', function (e) { var b = up(e.target, 'data-top', this); if (b) { S.top = b.getAttribute('data-top') === '1'; setOn($('topSeg'), 'data-top', S.top ? '1' : '0'); renderTable(); } });
    $('search').addEventListener('keyup', function () { S.q = this.value.replace(/^\s+|\s+$/g, ''); renderTable(); });
    $('resetBtn').addEventListener('click', function () { resetTable(); renderTeam(); });
    $('exportBtn').addEventListener('click', exportSheet);
    $('custHead').addEventListener('click', function (e) {
      var th = up(e.target, 'data-col', this); if (!th) return;
      var k = th.getAttribute('data-col');
      if (k === '월별') { if (S.hlCols[k]) delete S.hlCols[k]; else S.hlCols[k] = 1; renderTable(); return; }
      if (S.sort === k) S.dir = -S.dir; else { S.sort = k; S.dir = TEXT_COLS[k] ? 1 : -1; }
      S.hlSort = k; renderTable();
    });
    $('custBody').addEventListener('click', function (e) {
      var c = up(e.target, 'data-code', this); if (!c) return;
      var code = c.getAttribute('data-code');
      if (S.hlRows[code]) delete S.hlRows[code]; else S.hlRows[code] = 1;
      renderTable();
    });
    var goHandler = function (e) { var n = up(e.target, 'data-unit', this); if (n) goUnit(n.getAttribute('data-unit'), n.getAttribute('data-name')); };
    $('stack').addEventListener('click', goHandler);
    $('matrixBody').addEventListener('click', goHandler);
    $('bepBody').addEventListener('click', goHandler);
  }

  function start() { buildSlicer(); bind(); load(); }
  var started = false;
  function startOnce() { if (!started) { started = true; start(); } }
  if (window.Office) {
    Office.initialize = function () { startOnce(); };
    if (Office.onReady) Office.onReady(function () { startOnce(); });
  }
  window.__dashTest = { build: build, S: S, D: D, render: render };
})();
