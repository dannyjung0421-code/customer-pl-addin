/* 2026 거래처별 손익등급 - 브라우저용 실행 코드 (처음 대시보드와 같은 화면)
 * 데이터는 엑셀 매크로가 만든 화면 파일의 window.__CUSTPL_DATA 에 있다. 이 파일에는 데이터가 없다. */
(function () {
  var P = window.__CUSTPL_DATA;
  var root = document.getElementById('dash-root');
  if (!P || !root) { if (root) root.innerHTML = '<div class="full-msg">데이터가 없습니다. 엑셀의 [대시보드 열기] 버튼으로 여세요.</div>'; return; }
  var F = ['매출', '매출총이익', '변동비', '직접비', '판관비', '공헌이익', '직접이익', '영업이익'];
  function rowsOf(values) {
    var h = -1, i, j, out = [];
    if (!values || !values.length) return out;
    for (i = 0; i < Math.min(values.length, 10); i++) if (String(values[i][0]) === '기준') { h = i; break; }
    if (h < 0) return out;
    var head = values[h];
    for (i = h + 1; i < values.length; i++) {
      if (String(values[i][0]) !== '당월') continue;
      var o = {};
      for (j = 0; j < head.length; j++) o[String(head[j])] = values[i][j] === null ? '' : String(values[i][j]);
      out.push(o);
    }
    return out;
  }
  var pl = rowsOf(P.data), ex = rowsOf(P.exv), i, k, maxM = 1;
  for (i = 0; i < ex.length; i++) {
    ex[i]['사업부'] = '해당없음'; ex[i]['영업부'] = '해당없음';
    for (k = 0; k < F.length; k++) if (ex[i][F[k]] === undefined) ex[i][F[k]] = '0';
    pl.push(ex[i]);
  }
  for (i = 0; i < pl.length; i++) if (+pl[i]['월'] > maxM) maxM = +pl[i]['월'];
  var datasets = { pl: pl }, calcs = {}, cache = {}, fns = [];
  var DEFAULTS = { mode: 'C', m: maxM, tab: 'team', lv: '영업부서', team: 'FC1팀' };
  var params = {}; for (k in DEFAULTS) params[k] = DEFAULTS[k];
  function get(id) {
    if (id in datasets) return datasets[id];
    if (!(id in cache)) { var c = calcs[id]; cache[id] = c.fn.apply(null, c.inputs.map(get)); }
    return cache[id];
  }
  function run() { for (var n = 0; n < fns.length; n++) fns[n](); }
  window.dash = {
    calc: function (id, o) { calcs[id] = o; },
    loader: function () {},
    data: function (id) { return { status: 'ok', data: get(id), refreshing: false, meta: [] }; },
    onData: function (fn) { fns.push(fn); fn(); },
    params: function () { var o = {}; for (var x in params) o[x] = params[x]; return o; },
    setParams: function (o) { for (var x in o) params[x] = o[x]; run(); },
    resetParams: function () { params = {}; for (var x in DEFAULTS) params[x] = DEFAULTS[x]; run(); },
    setLink: function () {},
    refresh: function () {},
    colors: ['#3b6fd8', '#d9762b', '#2e9e8f', '#9b59b6', '#c4473b', '#5b8c2a', '#b5872a', '#64748b']
  };
  if (P.made) document.title = '2026 거래처별 손익등급 (' + P.made + ' 기준)';
  root.innerHTML = "<div class=\"lx\" id=\"lx\">\n  <header class=\"lx-head\">\n    <h2 class=\"lx-title\" id=\"title\">2026 거래처별 손익등급</h2>\n    <p class=\"lx-take dash-skeleton\" id=\"take\" data-source=\"rollup\">—</p>\n  </header>\n\n  <div class=\"lx-navrow\">\n  <nav class=\"lx-tabs\" role=\"tablist\" aria-label=\"화면\">\n    <button type=\"button\" role=\"tab\" data-tab=\"team\" id=\"tab-team\">조직별 거래처</button>\n    <button type=\"button\" role=\"tab\" data-tab=\"summary\" id=\"tab-summary\">요약</button>\n    <button type=\"button\" role=\"tab\" data-tab=\"rule\" id=\"tab-rule\">등급 기준</button>\n    <button type=\"button\" role=\"tab\" data-tab=\"excl\" id=\"tab-excl\">제외 영업부서</button>\n  </nav>\n    <div class=\"lx-bar lx-ctrl\">\n      <span class=\"lx-note hide-sm\" id=\"periodNote\"></span>\n      <div class=\"lx-seg\" role=\"group\" aria-label=\"집계 기준\" id=\"modeSeg\">\n        <button type=\"button\" data-mode=\"M\" id=\"modeM\">당월</button>\n        <button type=\"button\" data-mode=\"C\" id=\"modeC\">누적</button>\n      </div>\n      <select id=\"monthSel\" aria-label=\"기준 월\"></select>\n    </div>\n  </div>\n\n  <!-- 요약 -->\n  <section class=\"lx-page\" id=\"summary\">\n    <div class=\"lx-sumrow\">\n    <div class=\"lx-card\" id=\"sumKpiCard\">\n      <h3>분석대상 전체 손익 <span class=\"lx-note\" id=\"sumKpiNote\">TC·소매·소재·생두사업부 12개 영업부서</span></h3>\n      <div class=\"lx-kpis\" id=\"sumKpis\"></div>\n    </div>\n    <div class=\"lx-grades\" id=\"sumGrades\"></div>\n    </div>\n    <div class=\"lx-split\">\n    <div class=\"lx-card\" id=\"sumStackCard\">\n      <h3>사업부별 등급 구성</h3>\n      <div class=\"lx-bar\" style=\"justify-content: space-between\">\n        <p class=\"lx-note\">막대를 누르면 그 조직의 거래처 화면으로 이동합니다.</p>\n        <div class=\"lx-seg\" role=\"group\" aria-label=\"구성 기준\" id=\"basisSeg\">\n          <button type=\"button\" data-basis=\"sales\" id=\"basisSales\">매출 비중</button>\n          <button type=\"button\" data-basis=\"count\" id=\"basisCount\">거래처수 비중</button>\n        </div>\n      </div>\n      <div class=\"lx-legend\" id=\"legend1\"></div>\n      <div class=\"lx-stack\" id=\"sumStack\" data-source=\"matrix\"></div>\n    </div>\n    <div class=\"lx-card\" id=\"sumMatrixCard\">\n      <h3>사업부 · 영업부 · 영업부서별 등급 매트릭스</h3>\n      <p class=\"lx-note\">거래처수(매출 비중). 행을 누르면 그 조직의 거래처 화면으로 이동합니다. 금액 단위: 억원</p>\n      <div class=\"lx-tablewrap\">\n        <table id=\"matrix\" data-source=\"matrix\" data-row-key=\"key\">\n          <thead><tr>\n            <th class=\"l\" data-field=\"이름\">조직</th>\n            <th data-field=\"매출\">매출</th>\n            <th data-field=\"매출총이익\">매출총이익</th>\n            <th data-field=\"매출총이익률\">매출총이익률</th>\n            <th data-field=\"공헌이익률\">공헌이익률</th>\n            <th data-field=\"직접이익률\" class=\"hide-sm\">직접이익률</th>\n            <th data-field=\"영업이익률\">영업이익률</th>\n            <th data-field=\"거래처수\">거래처</th>\n            <th data-field=\"A거래처\">A</th>\n            <th data-field=\"B거래처\">B</th>\n            <th data-field=\"C거래처\">C</th>\n            <th data-field=\"D거래처\">D</th>\n          </tr></thead>\n          <tbody id=\"matrixBody\"></tbody>\n        </table>\n      </div>\n    </div>\n    </div>\n  </section>\n\n  <!-- 조직별 거래처 -->\n  <section class=\"lx-page\" id=\"team\" hidden>\n    <div class=\"lx-card lx-slicer\" id=\"teamSel\" role=\"group\" aria-label=\"조직 선택\"></div>\n    <div class=\"lx-sumrow\">\n      <div class=\"lx-card\" id=\"teamKpiCard\">\n        <h3><span id=\"teamKpiTitle\">영업부서 손익</span> <span class=\"lx-note\" id=\"teamPath\"></span></h3>\n        <div class=\"lx-kpis\" id=\"teamKpis\"></div>\n      </div>\n      <div class=\"lx-grades\" id=\"teamGrades\"></div>\n    </div>\n    <div class=\"lx-card\" id=\"teamTableCard\">\n      <div class=\"lx-bar\" style=\"justify-content: space-between\">\n        <h3>거래처별 손익</h3>\n        <div class=\"lx-chips\" id=\"gradeChips\" role=\"group\" aria-label=\"등급 필터\"></div>\n        <div class=\"lx-bar\" style=\"margin-left: auto\">\n          <div class=\"lx-seg\" role=\"group\" aria-label=\"표시 범위\" id=\"topSeg\">\n            <button type=\"button\" data-top=\"10\" id=\"top10\">매출 상위 10곳</button>\n            <button type=\"button\" data-top=\"all\" id=\"topAll\">전체</button>\n          </div>\n          <input type=\"search\" id=\"custSearch\" placeholder=\"거래처명·코드 검색\" aria-label=\"거래처 검색\">\n          <button type=\"button\" class=\"lx-clear\" id=\"csvBtn\">CSV로 저장</button>\n          <button type=\"button\" class=\"lx-clear\" id=\"hlClear\" hidden>필터 초기화</button>\n        </div>\n      </div>\n      <p class=\"lx-note\" id=\"teamTableNote\">금액 단위: 백만원. 월별 등급은 각 월 당월 기준입니다.</p>\n      <div class=\"lx-tablewrap\">\n        <table id=\"custTable\" data-source=\"grades\" data-row-key=\"key\">\n          <thead><tr id=\"custHead\">\n            <th class=\"l\" data-field=\"고객사명\" data-col=\"고객사명\"><span class=\"lx-th\"><button type=\"button\" class=\"lx-hl\" data-sortcol=\"고객사명\">거래처<span class=\"lx-arrow\" aria-hidden=\"true\"></span></button></span></th>\n            <th class=\"l lx-teamcol\" data-field=\"영업부서\" data-col=\"영업부서\"><span class=\"lx-th\"><button type=\"button\" class=\"lx-hl\" data-sortcol=\"영업부서\">영업부서<span class=\"lx-arrow\" aria-hidden=\"true\"></span></button></span></th>\n            <th data-field=\"등급\" data-col=\"등급\"><span class=\"lx-th\"><button type=\"button\" class=\"lx-hl\" data-sortcol=\"등급\">등급<span class=\"lx-arrow\" aria-hidden=\"true\"></span></button></span></th>\n            <th data-field=\"매출\" data-col=\"매출\"><span class=\"lx-th\"><button type=\"button\" class=\"lx-hl\" data-sortcol=\"매출\">매출<span class=\"lx-arrow\" aria-hidden=\"true\"></span></button></span></th>\n            <th data-field=\"매출총이익\" data-col=\"매출총이익\"><span class=\"lx-th\"><button type=\"button\" class=\"lx-hl\" data-sortcol=\"매출총이익\">매출총이익<span class=\"lx-arrow\" aria-hidden=\"true\"></span></button></span></th>\n            <th data-field=\"매출총이익률\" data-col=\"매출총이익률\"><span class=\"lx-th\"><button type=\"button\" class=\"lx-hl\" data-sortcol=\"매출총이익률\">매출총이익률<span class=\"lx-arrow\" aria-hidden=\"true\"></span></button></span></th>\n            <th data-field=\"공헌이익\" data-col=\"공헌이익\"><span class=\"lx-th\"><button type=\"button\" class=\"lx-hl\" data-sortcol=\"공헌이익\">공헌이익<span class=\"lx-arrow\" aria-hidden=\"true\"></span></button></span></th>\n            <th data-field=\"공헌이익률\" data-col=\"공헌이익률\"><span class=\"lx-th\"><button type=\"button\" class=\"lx-hl\" data-sortcol=\"공헌이익률\">공헌이익률<span class=\"lx-arrow\" aria-hidden=\"true\"></span></button></span></th>\n            <th class=\"hide-sm\" data-field=\"직접이익\" data-col=\"직접이익\"><span class=\"lx-th\"><button type=\"button\" class=\"lx-hl\" data-sortcol=\"직접이익\">직접이익<span class=\"lx-arrow\" aria-hidden=\"true\"></span></button></span></th>\n            <th class=\"hide-sm\" data-field=\"직접이익률\" data-col=\"직접이익률\"><span class=\"lx-th\"><button type=\"button\" class=\"lx-hl\" data-sortcol=\"직접이익률\">직접이익률<span class=\"lx-arrow\" aria-hidden=\"true\"></span></button></span></th>\n            <th class=\"hide-sm\" data-field=\"영업이익\" data-col=\"영업이익\"><span class=\"lx-th\"><button type=\"button\" class=\"lx-hl\" data-sortcol=\"영업이익\">영업이익<span class=\"lx-arrow\" aria-hidden=\"true\"></span></button></span></th>\n            <th class=\"hide-sm\" data-field=\"영업이익률\" data-col=\"영업이익률\"><span class=\"lx-th\"><button type=\"button\" class=\"lx-hl\" data-sortcol=\"영업이익률\">영업이익률<span class=\"lx-arrow\" aria-hidden=\"true\"></span></button></span></th>\n            <th class=\"l hide-sm\" data-col=\"월별\"><span class=\"lx-th\"><button type=\"button\" class=\"lx-hl\" data-hl=\"월별\">월별 등급 (1→9월)</button></span></th>\n          </tr></thead>\n          <tbody id=\"custBody\"></tbody>\n        </table>\n      </div>\n    </div>\n  </section>\n\n  <!-- 등급 기준 -->\n  <section class=\"lx-page\" id=\"rule\" hidden>\n    <div class=\"lx-split lx-split-rule\">\n    <div class=\"lx-card\" id=\"ruleDefCard\">\n      <h3>등급은 이렇게 나눕니다</h3>\n      <p class=\"lx-note\">거래처의 공헌이익률을 소속 영업부서의 기준선 3개와 비교합니다. 기준선은 영업부서 전체 손익 구조에서 계산하며, 당월은 그 달, 누적은 1월부터 기준 월까지의 합계로 다시 계산합니다.</p>\n      <dl class=\"lx-def\">\n        <dt><span class=\"lx-pill\" style=\"--gc: var(--g-a)\">A</span> 목표달성</dt><dd>공헌이익률 ≥ 목표 공헌이익률 (영업이익 BEP + 5%p, 즉 영업이익률 5% 수준)</dd>\n        <dt><span class=\"lx-pill\" style=\"--gc: var(--g-b)\">B</span> 영업이익 BEP</dt><dd>공헌이익률 ≥ 영업이익 BEP 공헌이익률 = (판관비 − 변동비) ÷ 매출. 고정비와 배부비까지 회수하는 수준</dd>\n        <dt><span class=\"lx-pill\" style=\"--gc: var(--g-c)\">C</span> 직접영업이익 BEP</dt><dd>공헌이익률 ≥ 직접영업이익 BEP 공헌이익률 = (직접비 − 변동비) ÷ 매출. 거래처 직접고정비까지만 회수하는 수준</dd>\n        <dt><span class=\"lx-pill\" style=\"--gc: var(--g-d)\">D</span> 직접영업이익 적자</dt><dd>공헌이익률이 직접영업이익 BEP에 못 미침. 거래를 유지할수록 직접영업이익이 줄어드는 구간</dd>\n        <dt><span class=\"lx-pill\" style=\"--gc: var(--g-x)\">-</span> 등급외</dt><dd>해당 기간 매출이 0 이하인 거래처 (반품·조정만 있는 경우). 등급 집계에서 별도로 셉니다</dd>\n      </dl>\n      <dl class=\"lx-def\">\n        <dt>직접비</dt><dd>[판관] 인건비 ~ 기타판관비 합계</dd>\n        <dt>간접/배부비</dt><dd>[판관] OP1 ~ OP3배부비 합계</dd>\n        <dt>변동비</dt><dd>운반비(직접·간접·대상·고양) + 광고선전비(직접·간접) + 판매촉진비 + 판매수수료(직접·간접)</dd>\n        <dt>공헌이익</dt><dd>매출총이익 − 변동비</dd>\n        <dt>직접영업이익</dt><dd>매출총이익 − 직접비</dd>\n        <dt>영업이익</dt><dd>매출총이익 − 판관비</dd>\n      </dl>\n      <p class=\"lx-note\">원본 RAW의 각 열 값을 그대로 사용합니다. </p>\n    </div>\n    <div class=\"lx-card\" id=\"ruleTableCard\">\n      <h3 id=\"ruleTableTitle\">영업부서별 기준선</h3>\n      <div class=\"lx-tablewrap\">\n        <table id=\"bepTable\" data-source=\"bep\" data-row-key=\"key\">\n          <thead><tr>\n            <th class=\"l\" data-field=\"사업부\">사업부</th>\n            <th class=\"l\" data-field=\"영업부서\">영업부서</th>\n            <th data-field=\"매출\">매출(억원)</th>\n            <th data-field=\"공헌이익률\">실적 공헌이익률</th>\n            <th data-field=\"목표\">A 목표</th>\n            <th data-field=\"영업BEP\">B 영업이익 BEP</th>\n            <th data-field=\"직접BEP\">C 직접영업이익 BEP</th>\n          </tr></thead>\n          <tbody id=\"bepBody\"></tbody>\n        </table>\n      </div>\n    </div>\n    </div>\n  </section>\n\n  <!-- 제외 -->\n  <section class=\"lx-page\" id=\"excl\" hidden>\n    <div class=\"lx-card\" id=\"exclCard\">\n      <h3>거래처별 손익 대상이 아닌 영업부서</h3>\n      <p class=\"lx-note\">사업부 매핑에 없는 영업부서입니다. 등급 집계와 요약 수치에서 빠져 있습니다. 금액 단위: 억원</p>\n      <div class=\"lx-tablewrap\">\n        <table id=\"exclTable\" data-source=\"excluded\" data-row-key=\"key\">\n          <thead><tr>\n            <th class=\"l\" data-field=\"영업부서\">영업부서</th>\n            <th data-field=\"거래처수\">거래처</th>\n            <th data-field=\"매출\">매출</th>\n            <th data-field=\"매출총이익\">매출총이익</th>\n            <th data-field=\"공헌이익\">공헌이익</th>\n            <th data-field=\"직접이익\">직접영업이익</th>\n            <th data-field=\"영업이익\">영업이익</th>\n            <th data-field=\"영업이익률\">영업이익률</th>\n          </tr></thead>\n          <tbody id=\"exclBody\"></tbody>\n        </table>\n      </div>\n    </div>\n  </section>\n</div>";
  var t; window.addEventListener('resize', function () { clearTimeout(t); t = setTimeout(run, 150); });
})();

if (window.dash) {

(() => {
  const ORDER = [
    ['TC사업부', [['FC영업부', ['FC1팀', 'FC2팀']], ['CC영업부', ['뉴채널팀', 'CNT플랫폼팀', 'CNT영업팀', '머신운영팀']]]],
    ['소매사업부', [['소매사업부', ['KAM1팀', 'KAM2팀']]]],
    ['소재사업부', [['소재사업부', ['소재영업1팀', '소재영업2팀', '소재사업부(영업)']]]],
    ['생두사업부', [['생두사업부', ['생두운영팀']]]],
  ];
  const GRADES = ['A', 'B', 'C', 'D'];
  const GNAME = { A: '목표달성', B: '영업이익 BEP', C: '직접영업이익 BEP', D: '직접영업이익 적자', '등급외': '매출 없음' };
  const GVAR = { A: 'var(--g-a)', B: 'var(--g-b)', C: 'var(--g-c)', D: 'var(--g-d)', '등급외': 'var(--g-x)' };

  /* ---------- calculations ---------- */
  dash.calc('bep', {
    inputs: ['pl'],
    title: '영업부서별 등급 기준선',
    description: '영업부서 합계로 영업이익 BEP=(판관비−변동비)/매출, 직접영업이익 BEP=(직접비−변동비)/매출, 목표=영업이익 BEP+5%p를 당월·누적별로 계산',
    fn: (pl) => {
      const TARGET_MARGIN = 0.05;
      const agg = new Map();
      for (const r of pl) {
        if (r['사업부'] === '해당없음') continue;
        const m = +r['월'];
        const targets = [['M', m]];
        for (let k = m; k <= 9; k++) targets.push(['C', k]);
        for (const [mode, k] of targets) {
          const key = mode + '|' + k + '|' + r['영업부서'];
          let a = agg.get(key);
          if (!a) { a = { key, 기준: mode, 월: k, 사업부: r['사업부'], 영업부: r['영업부'], 영업부서: r['영업부서'], 매출: 0, 변동비: 0, 직접비: 0, 판관비: 0, 공헌이익: 0 }; agg.set(key, a); }
          a.매출 += +r['매출']; a.변동비 += +r['변동비']; a.직접비 += +r['직접비']; a.판관비 += +r['판관비']; a.공헌이익 += +r['공헌이익'];
        }
      }
      return [...agg.values()].map((a) => {
        const 영업BEP = a.매출 ? (a.판관비 - a.변동비) / a.매출 : null;
        const 직접BEP = a.매출 ? (a.직접비 - a.변동비) / a.매출 : null;
        return { ...a, 공헌이익률: a.매출 ? a.공헌이익 / a.매출 : null, 목표: 영업BEP == null ? null : 영업BEP + TARGET_MARGIN, 영업BEP, 직접BEP };
      });
    },
  });

  dash.calc('grades', {
    inputs: ['pl', 'bep'],
    title: '거래처별 손익등급',
    description: '거래처(영업부서별)의 당월·누적 손익과 공헌이익률을 소속 영업부서 기준선과 비교해 A/B/C/D 등급 부여 (매출 0 이하는 등급외)',
    fn: (pl, bep) => {
      const F = ['매출', '매출총이익', '변동비', '직접비', '판관비', '공헌이익', '직접이익', '영업이익'];
      const th = new Map(bep.map((b) => [b.key, b]));
      const cust = new Map();
      for (const r of pl) {
        if (r['사업부'] === '해당없음') continue;
        const id = r['영업부서'] + '|' + r['고객코드'];
        let c = cust.get(id);
        if (!c) { c = { 사업부: r['사업부'], 영업부: r['영업부'], 영업부서: r['영업부서'], 고객코드: String(r['고객코드']), 고객사명: r['고객사명'], mon: {} }; cust.set(id, c); }
        const m = +r['월'];
        const v = c.mon[m] || (c.mon[m] = F.map(() => 0));
        F.forEach((f, i) => { v[i] += +r[f]; });
      }
      const grade = (v, t) => {
        if (!t || v[0] <= 0) return '등급외';
        const rate = v[5] / v[0];
        return rate >= t.목표 ? 'A' : rate >= t.영업BEP ? 'B' : rate >= t.직접BEP ? 'C' : 'D';
      };
      const out = [];
      for (const c of cust.values()) {
        const mg = {};
        for (let k = 1; k <= 9; k++) mg[k] = c.mon[k] ? grade(c.mon[k], th.get('M|' + k + '|' + c.영업부서)) : '';
        const row = (mode, k, v, g) => {
          const o = { key: mode + '|' + k + '|' + c.영업부서 + '|' + c.고객코드, 기준: mode, 월: k, 사업부: c.사업부, 영업부: c.영업부, 영업부서: c.영업부서, 고객코드: c.고객코드, 고객사명: c.고객사명,
            매출: v[0], 매출총이익: v[1], 공헌이익: v[5], 직접이익: v[6], 영업이익: v[7],
            매출총이익률: v[0] > 0 ? v[1] / v[0] : null, 공헌이익률: v[0] > 0 ? v[5] / v[0] : null, 직접이익률: v[0] > 0 ? v[6] / v[0] : null, 영업이익률: v[0] > 0 ? v[7] / v[0] : null, 등급: g };
          for (let j = 1; j <= 9; j++) o[j + '월'] = mg[j];
          return o;
        };
        const cum = F.map(() => 0);
        let seen = false;
        for (let k = 1; k <= 9; k++) {
          if (c.mon[k]) { c.mon[k].forEach((x, i) => { cum[i] += x; }); seen = true; out.push(row('M', k, c.mon[k], mg[k])); }
          if (seen) out.push(row('C', k, cum.slice(), grade(cum, th.get('C|' + k + '|' + c.영업부서))));
        }
      }
      return out;
    },
  });

  dash.calc('rollup', {
    inputs: ['grades'],
    title: '조직·등급별 집계',
    description: '전사·사업부·영업부·영업부서 단위로 등급별 거래처수와 손익 합계, 이익률, 매출 비중을 집계',
    fn: (grades) => {
      const agg = new Map();
      for (const r of grades) {
        const units = [['전사', '전사'], ['사업부', r.사업부], ['영업부', r.영업부], ['영업부서', r.영업부서]];
        for (const [u, n] of units) {
          for (const g of [r.등급, '전체']) {
            const key = r.기준 + '|' + r.월 + '|' + u + '|' + n + '|' + g;
            let a = agg.get(key);
            if (!a) { a = { key, 기준: r.기준, 월: r.월, 단위: u, 이름: n, 등급: g, 거래처수: 0, 매출: 0, 매출총이익: 0, 공헌이익: 0, 직접이익: 0, 영업이익: 0 }; agg.set(key, a); }
            a.거래처수 += 1; a.매출 += r.매출; a.매출총이익 += r.매출총이익; a.공헌이익 += r.공헌이익; a.직접이익 += r.직접이익; a.영업이익 += r.영업이익;
          }
        }
      }
      const rows = [...agg.values()];
      for (const a of rows) {
        const tot = agg.get(a.기준 + '|' + a.월 + '|' + a.단위 + '|' + a.이름 + '|전체');
        a.매출총이익률 = a.매출 ? a.매출총이익 / a.매출 : null;
        a.공헌이익률 = a.매출 ? a.공헌이익 / a.매출 : null;
        a.직접이익률 = a.매출 ? a.직접이익 / a.매출 : null;
        a.영업이익률 = a.매출 ? a.영업이익 / a.매출 : null;
        a.매출비중 = tot.매출 ? a.매출 / tot.매출 : null;
        a.거래처비중 = tot.거래처수 ? a.거래처수 / tot.거래처수 : null;
      }
      return rows;
    },
  });

  dash.calc('matrix', {
    inputs: ['rollup'],
    title: '조직별 등급 매트릭스',
    description: '조직 단위 한 행에 등급별 거래처수·매출·매출 비중을 펼친 표',
    fn: (rollup) => {
      const m = new Map();
      for (const r of rollup) {
        const key = r.기준 + '|' + r.월 + '|' + r.단위 + '|' + r.이름;
        let o = m.get(key);
        if (!o) { o = { key, 기준: r.기준, 월: r.월, 단위: r.단위, 이름: r.이름 }; for (const g of ['A', 'B', 'C', 'D', '등급외']) { o[g + '거래처'] = 0; o[g + '매출'] = 0; o[g + '비중'] = 0; o[g + '거래처비중'] = 0; } m.set(key, o); }
        if (r.등급 === '전체') Object.assign(o, { 거래처수: r.거래처수, 매출: r.매출, 매출총이익: r.매출총이익, 매출총이익률: r.매출총이익률, 공헌이익: r.공헌이익, 직접이익: r.직접이익, 영업이익: r.영업이익, 공헌이익률: r.공헌이익률, 직접이익률: r.직접이익률, 영업이익률: r.영업이익률 });
        else { o[r.등급 + '거래처'] = r.거래처수; o[r.등급 + '매출'] = r.매출; o[r.등급 + '비중'] = r.매출비중; o[r.등급 + '거래처비중'] = r.거래처비중; }
      }
      return [...m.values()];
    },
  });

  dash.calc('excluded', {
    inputs: ['pl'],
    title: '제외 영업부서 손익',
    description: '사업부 매핑에 없는 영업부서의 당월·누적 손익 합계',
    fn: (pl) => {
      const agg = new Map();
      for (const r of pl) {
        if (r['사업부'] !== '해당없음') continue;
        const m = +r['월'];
        const targets = [['M', m]];
        for (let k = m; k <= 9; k++) targets.push(['C', k]);
        for (const [mode, k] of targets) {
          const key = mode + '|' + k + '|' + r['영업부서'];
          let a = agg.get(key);
          if (!a) { a = { key, 기준: mode, 월: k, 영업부서: r['영업부서'], 거래처: new Set(), 매출: 0, 매출총이익: 0, 공헌이익: 0, 직접이익: 0, 영업이익: 0 }; agg.set(key, a); }
          a.거래처.add(r['고객코드']);
          a.매출 += +r['매출']; a.매출총이익 += +r['매출총이익']; a.공헌이익 += +r['공헌이익']; a.직접이익 += +r['직접이익']; a.영업이익 += +r['영업이익'];
        }
      }
      return [...agg.values()].map(({ 거래처, ...a }) => ({ ...a, 거래처수: 거래처.size, 영업이익률: a.매출 ? a.영업이익 / a.매출 : null }));
    },
  });


  /* ---------- helpers ---------- */
  const $ = (id) => document.getElementById(id);
  const root = $('lx');
  root.style.setProperty('--g-b', dash.colors[0]);
  const fEok = (v) => v == null ? '—' : d3.format(',.1f')(v / 1e8);
  const fMil = (v) => v == null ? '—' : d3.format(',.0f')(v / 1e6);
  const fPct = (v) => v == null || !isFinite(v) ? '—' : d3.format('.1%')(v);
  const fInt = d3.format(',');
  const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; };
  const negCls = (v) => (v != null && v < 0 ? 'neg' : '');
  const state = { basis: 'sales', gf: '전체', q: '', sort: '매출', dir: -1, top: '10', sorted: false, hlSort: null, hlRows: new Set(), hlCols: new Set() };

  const P = () => { const p = dash.params(); const tab = p.tab === 'bu' ? 'team' : (p.tab || 'team'); return { mode: p.mode || 'C', m: +p.m || 9, tab, lv: p.lv || '영업부서', team: p.team || 'FC1팀' }; };
  const goUnit = (lv, team) => { dash.setParams({ tab: 'team', lv, team }); dash.setLink('team'); };
  const periodLabel = (p) => p.mode === 'M' ? `2026년 ${p.m}월 당월` : `2026년 1~${p.m}월 누적`;

  // tooltip
  const tip = el('div', 'lx-tip'); tip.hidden = true;
  const showTip = (host, ev, lines) => {
    if (tip.parentNode !== host) host.appendChild(tip);
    tip.replaceChildren(...lines.map((l, i) => { const d = el('div', i === 0 ? 't' : '', l); return d; }));
    tip.hidden = false;
    const hb = host.getBoundingClientRect();
    let x = ev.clientX - hb.left + 12, y = ev.clientY - hb.top + 12;
    const tw = tip.offsetWidth, thh = tip.offsetHeight;
    if (x + tw > hb.width) x = Math.max(0, ev.clientX - hb.left - tw - 12);
    if (y + thh > hb.height + 40) y = ev.clientY - hb.top - thh - 12;
    tip.style.left = x + 'px'; tip.style.top = y + 'px';
  };
  const hideTip = () => { tip.hidden = true; };

  // controls
  const monthSel = $('monthSel');
  for (let k = 1; k <= 9; k++) monthSel.appendChild(new Option(k + '월', String(k)));
  monthSel.addEventListener('change', () => dash.setParams({ m: +monthSel.value }));
  $('modeSeg').addEventListener('click', (e) => { const b = e.target.closest('button[data-mode]'); if (b) dash.setParams({ mode: b.dataset.mode }); });
  document.querySelector('.lx-tabs').addEventListener('click', (e) => { const b = e.target.closest('button[data-tab]'); if (b) { dash.setParams({ tab: b.dataset.tab }); dash.setLink(b.dataset.tab); } });
  $('basisSeg').addEventListener('click', (e) => { const b = e.target.closest('button[data-basis]'); if (b) { state.basis = b.dataset.basis; render(); } });
  $('teamSel').addEventListener('click', (e) => { const b = e.target.closest('button[data-team]'); if (b) dash.setParams({ lv: b.dataset.lv, team: b.dataset.team }); });
  $('custSearch').addEventListener('input', (e) => { state.q = e.target.value.trim().toLowerCase(); renderTeamTable(); });
  const toggle = (set, v) => { if (set.has(v)) set.delete(v); else set.add(v); };
  $('custHead').addEventListener('click', (e) => {
    const s = e.target.closest('button[data-sortcol]');
    if (s) {
      const k = s.dataset.sortcol;
      if (state.sort === k) state.dir *= -1; else { state.sort = k; state.dir = ['고객사명', '영업부서', '등급'].includes(k) ? 1 : -1; }
      state.sorted = true; state.hlSort = k;
      renderTeamTable(); return;
    }
    const h = e.target.closest('button[data-hl]');
    if (h) { toggle(state.hlCols, h.dataset.hl); renderTeamTable(); }
  });
  $('custBody').addEventListener('click', (e) => { const td = e.target.closest('td.namecell'); if (td) { toggle(state.hlRows, td.dataset.code); renderTeamTable(); } });
  $('topSeg').addEventListener('click', (e) => { const b = e.target.closest('button[data-top]'); if (b) { state.top = b.dataset.top; renderTeamTable(); } });
  $('hlClear').addEventListener('click', () => { Object.assign(state, { gf: '전체', q: '', sort: '매출', dir: -1, top: '10', sorted: false, hlSort: null }); state.hlRows.clear(); state.hlCols.clear(); $('custSearch').value = ''; renderTeam(P()); });

  // team select options
  const teamSel = $('teamSel');
  for (const [bu, subs] of ORDER) {
    const mkb = (cls, txt, lv) => { const b = el('button', cls, txt); b.type = 'button'; b.dataset.team = txt; b.dataset.lv = lv; return b; };
    const grp = el('div', 'grp'); grp.appendChild(mkb('gl', bu, '사업부'));
    const box = el('div', 'lx-chips');
    for (const [sub, teams] of subs) {
      const w = el('span', 'sub');
      if (sub !== bu) w.appendChild(mkb('sl', sub, '영업부'));
      for (const t of teams) w.appendChild(mkb(null, t, '영업부서'));
      box.appendChild(w);
    }
    grp.appendChild(box); teamSel.appendChild(grp);
  }
  const teamPath = (t) => { for (const [bu, subs] of ORDER) for (const [sub, teams] of subs) if (teams.includes(t)) return bu === sub ? bu : bu + ' › ' + sub; return ''; };

  const legend = (host) => host.replaceChildren(...GRADES.map((g) => { const s = el('span'); const i = el('i'); i.style.setProperty('--sw', GVAR[g]); s.append(i, document.createTextNode(g + ' ' + GNAME[g])); return s; }));
  legend($('legend1'));

  const failed = (...ids) => ids.some((id) => ['error', 'connect', 'declined', 'missing'].includes(dash.data(id).status));
  const ready = (...ids) => ids.every((id) => dash.data(id).status === 'ok');

  /* ---------- pieces ---------- */
  function kpis(host, unit, name, p, mil) {
    const fm = (v) => mil ? fMil(v) + '백만' : fEok(v) + '억';
    const key = `기준=${p.mode}&월=${p.m}&단위=${encodeURIComponent(unit)}&이름=${encodeURIComponent(name)}&등급=${encodeURIComponent('전체')}`;
    const rows = dash.data('rollup').data || [];
    const r = rows.find((x) => x.기준 === p.mode && x.월 === p.m && x.단위 === unit && x.이름 === name && x.등급 === '전체');
    const items = [
      ['매출', '매출', (x) => fm(x.매출), null],
      ['공헌이익', '공헌이익', (x) => fm(x.공헌이익), (x) => '공헌이익률 ' + fPct(x.공헌이익률)],
      ['직접영업이익', '직접이익', (x) => fm(x.직접이익), (x) => '직접영업이익률 ' + fPct(x.직접이익률)],
      ['영업이익', '영업이익', (x) => fm(x.영업이익), (x) => '영업이익률 ' + fPct(x.영업이익률)],
      ['거래처', '거래처수', (x) => fInt(x.거래처수) + '곳', null],
    ];
    host.replaceChildren(...items.map(([lab, field, f, sub]) => {
      const box = el('div', 'lx-kpi');
      box.appendChild(el('span', 'lab', lab));
      const v = el('span', 'val' + (r ? '' : ' dash-skeleton'), r ? f(r) : '—');
      if (r && field !== '거래처수' && r[field] < 0) v.classList.add('neg');
      v.dataset.source = 'rollup'; v.dataset.field = field; v.setAttribute('data-where', key);
      box.appendChild(v);
      if (sub) box.appendChild(el('span', 'sub', r ? sub(r) : ''));
      return box;
    }));
  }

  function gradeCards(host, unit, name, p) {
    const rows = dash.data('rollup').data || [];
    host.replaceChildren(...GRADES.map((g) => {
      const r = rows.find((x) => x.기준 === p.mode && x.월 === p.m && x.단위 === unit && x.이름 === name && x.등급 === g);
      const where = `기준=${p.mode}&월=${p.m}&단위=${encodeURIComponent(unit)}&이름=${encodeURIComponent(name)}&등급=${g}`;
      const c = el('div', 'lx-card lx-grade'); c.style.setProperty('--gc', GVAR[g]); c.id = `${host.id}-${g}`;
      const head = el('div', 'gname'); head.append(el('span', 'gletter', g), el('span', 'gdesc', GNAME[g])); c.appendChild(head);
      const dl = el('dl');
      const add = (dt, field, text, big) => { dl.appendChild(el('dt', null, dt)); const dd = el('dd', big ? 'big' : '', text); if (r) { dd.dataset.source = 'rollup'; dd.dataset.field = field; dd.setAttribute('data-where', where); } if (r && /이익/.test(field) && r[field] < 0) dd.classList.add('neg'); dl.appendChild(dd); };
      add('거래처', '거래처수', r ? fInt(r.거래처수) + '곳' : '0곳', true);
      add('매출', '매출', r ? fEok(r.매출) + '억' : '—');
      add('매출 비중', '매출비중', r ? fPct(r.매출비중) : '—');
      add('공헌이익률', '공헌이익률', r ? fPct(r.공헌이익률) : '—');
      add('직접영업이익률', '직접이익률', r ? fPct(r.직접이익률) : '—');
      c.appendChild(dl);
      return c;
    }));
  }

  function stackRows(host, items, p, onClick) {
    const rows = dash.data('matrix').data || [];
    const basis = state.basis;
    host.replaceChildren(...items.map(({ unit, name, label, cls }) => {
      const r = rows.find((x) => x.기준 === p.mode && x.월 === p.m && x.단위 === unit && x.이름 === name);
      const row = el('div', 'lx-srow');
      row.appendChild(el('span', 'nm ' + (cls || ''), label));
      const bar = el('div', 'lx-sbar'); bar.id = `${host.id}-${unit}-${name}`;
      if (r) {
        for (const g of GRADES) {
          const share = basis === 'sales' ? r[g + '비중'] : r[g + '거래처비중'];
          if (!share || share <= 0) continue;
          const seg = el('div'); seg.style.width = (share * 100) + '%'; seg.style.background = GVAR[g];
          if (share >= 0.07) seg.textContent = g + ' ' + d3.format('.0%')(share);
          seg.addEventListener('mousemove', (ev) => showTip(host, ev, [`${label} · ${g} ${GNAME[g]}`, `거래처 ${fInt(r[g + '거래처'])}곳`, `매출 ${fEok(r[g + '매출'])}억`, `매출 비중 ${fPct(r[g + '비중'])} · 거래처 비중 ${fPct(r[g + '거래처비중'])}`]));
          seg.addEventListener('mouseleave', hideTip);
          bar.appendChild(seg);
        }
        if (onClick) bar.addEventListener('click', () => onClick(unit, name));
      }
      row.appendChild(bar);
      return row;
    }));
  }

  /* ---------- tabs ---------- */
  function renderSummary(p) {
    kpis($('sumKpis'), '전사', '전사', p);
    gradeCards($('sumGrades'), '전사', '전사', p);
    const items = [];
    for (const [bu, subs] of ORDER) {
      items.push({ unit: '사업부', name: bu, label: bu });
      if (subs.length > 1) for (const [sub] of subs) items.push({ unit: '영업부', name: sub, label: sub, cls: 'child' });
    }
    items.push({ unit: '전사', name: '전사', label: '분석대상 전체', cls: 'total' });
    if (failed('matrix')) $('sumStack').replaceChildren(el('div', 'lx-fail', 'Failed to load'));
    else stackRows($('sumStack'), items, p, (unit, name) => { if (unit !== '전사') goUnit(unit, name); });

    const rows = dash.data('matrix').data || [];
    const body = $('matrixBody');
    const trs = [];
    const mk = (unit, name, label, lv, team) => {
      const r = rows.find((x) => x.기준 === p.mode && x.월 === p.m && x.단위 === unit && x.이름 === name);
      if (!r) return;
      const tr = el('tr', 'lv' + lv + (unit !== '전사' ? ' click' : '') + (unit === '전사' ? ' tot' : '')); tr.dataset.key = r.key;
      tr.appendChild(el('td', 'l', label));
      tr.appendChild(el('td', negCls(r.매출), fEok(r.매출)));
      tr.appendChild(el('td', negCls(r.매출총이익), fEok(r.매출총이익)));
      tr.appendChild(el('td', negCls(r.매출총이익률), fPct(r.매출총이익률)));
      tr.appendChild(el('td', negCls(r.공헌이익률), fPct(r.공헌이익률)));
      tr.appendChild(el('td', 'hide-sm ' + negCls(r.직접이익률), fPct(r.직접이익률)));
      tr.appendChild(el('td', negCls(r.영업이익률), fPct(r.영업이익률)));
      tr.appendChild(el('td', null, fInt(r.거래처수)));
      for (const g of GRADES) tr.appendChild(el('td', null, `${fInt(r[g + '거래처'])} (${d3.format('.0%')(r[g + '비중'] || 0)})`));
      if (unit !== '전사') tr.addEventListener('click', () => goUnit(unit, name));
      trs.push(tr);
    };
    for (const [bu, subs] of ORDER) {
      mk('사업부', bu, bu, 0);
      for (const [sub, teams] of subs) {
        const nested = sub !== bu;
        if (nested) mk('영업부', sub, sub, 1);
        for (const t of teams) mk('영업부서', t, t, nested ? 2 : 1, t);
      }
    }
    mk('전사', '전사', '분석대상 전체', 0);
    body.replaceChildren(...trs);
    if (failed('matrix')) { const tr = el('tr'); const td = el('td', 'lx-fail', 'Failed to load'); td.colSpan = 12; tr.appendChild(td); body.replaceChildren(tr); }
  }

  function renderTeam(p) {
    teamSel.querySelectorAll('button[data-team]').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.team === p.team && b.dataset.lv === p.lv)));
    $('teamPath').textContent = p.lv === '영업부서' ? teamPath(p.team) : p.lv === '영업부' ? 'TC사업부' : '';
    $('teamKpiTitle').textContent = p.team + ' 손익 · ' + periodLabel(p);
    kpis($('teamKpis'), p.lv, p.team, p, true);
    gradeCards($('teamGrades'), p.lv, p.team, p);
    const gc = $('gradeChips');
    gc.replaceChildren(...['전체', 'AB', 'CD', ...GRADES, '등급외'].map((g) => { const b = el('button', null, g === '전체' ? '전체 등급' : g === '등급외' ? '등급외' : g === 'AB' ? 'A+B' : g === 'CD' ? 'C+D' : g + ' ' + GNAME[g]); b.type = 'button'; b.setAttribute('aria-pressed', String(state.gf === g)); b.addEventListener('click', () => { state.gf = g; renderTeam(P()); }); return b; }));
    renderTeamTable();
  }

  function renderTeamTable() {
    const p = P();
    const body = $('custBody');
    if (failed('grades')) { const tr = el('tr'); const td = el('td', 'lx-fail', 'Failed to load'); td.colSpan = 13; tr.appendChild(td); body.replaceChildren(tr); return; }
    const fld = p.lv;
    $('custTable').classList.toggle('one-team', p.lv === '영업부서');
    let rows = (dash.data('grades').data || []).filter((r) => r.기준 === p.mode && r.월 === p.m && r[fld] === p.team);
    if (state.gf !== '전체') rows = rows.filter((r) => state.gf.length === 2 ? state.gf.includes(r.등급) : r.등급 === state.gf);
    if (state.q) rows = rows.filter((r) => (r.고객사명 + ' ' + r.고객코드).toLowerCase().includes(state.q));
    const total = rows.length;
    if (state.top === '10') rows = rows.slice().sort((a, b) => b.매출 - a.매출).slice(0, 10);
    const k = state.sort, dir = state.dir;
    rows.sort((a, b) => { const va = a[k], vb = b[k]; if (va == null) return 1; if (vb == null) return -1; return (typeof va === 'string' ? va.localeCompare(vb, 'ko') : va - vb) * dir; });
    window.__custRows = rows; window.__custInfo = { team: p.team, label: periodLabel(p), mode: p.mode, m: p.m };
    $('top10').setAttribute('aria-pressed', String(state.top === '10'));
    $('topAll').setAttribute('aria-pressed', String(state.top !== '10'));
    const colOn = (c) => state.hlCols.has(c) || state.hlSort === c;
    $('hlClear').hidden = !(state.hlRows.size || state.hlCols.size || state.hlSort || state.gf !== '전체' || state.q || state.top !== '10' || state.sort !== '매출' || state.dir !== -1);
    document.querySelectorAll('#custHead th').forEach((th) => {
      const col = th.dataset.col;
      th.classList.toggle('hl1', colOn(col));
      th.setAttribute('aria-sort', col === k ? (dir > 0 ? 'ascending' : 'descending') : 'none');
      const a = th.querySelector('.lx-arrow'); if (a) a.textContent = col === k ? (dir > 0 ? '△' : '▽') : '';
    });
    const COLS = ['고객사명', '영업부서', '등급', '매출', '매출총이익', '매출총이익률', '공헌이익', '공헌이익률', '직접이익', '직접이익률', '영업이익', '영업이익률', '월별'];
    body.replaceChildren(...rows.map((r) => {
      const tr = el('tr'); tr.dataset.key = r.key;
      const name = el('td', 'l namecell'); name.textContent = r.고객사명; name.title = r.고객코드 + ' · 누르면 행 음영'; name.dataset.code = r.고객코드; tr.appendChild(name);
      const tc = el('td', 'l lx-teamcol', r.영업부서); tr.appendChild(tc);
      const g = el('td'); const pill = el('span', 'lx-pill', r.등급 === '등급외' ? '-' : r.등급); pill.style.setProperty('--gc', GVAR[r.등급]); g.appendChild(pill); tr.appendChild(g);
      tr.appendChild(el('td', negCls(r.매출), fMil(r.매출)));
      tr.appendChild(el('td', negCls(r.매출총이익), fMil(r.매출총이익)));
      tr.appendChild(el('td', negCls(r.매출총이익률), fPct(r.매출총이익률)));
      tr.appendChild(el('td', negCls(r.공헌이익), fMil(r.공헌이익)));
      tr.appendChild(el('td', negCls(r.공헌이익률), fPct(r.공헌이익률)));
      tr.appendChild(el('td', 'hide-sm ' + negCls(r.직접이익), fMil(r.직접이익)));
      tr.appendChild(el('td', 'hide-sm ' + negCls(r.직접이익률), fPct(r.직접이익률)));
      tr.appendChild(el('td', 'hide-sm ' + negCls(r.영업이익), fMil(r.영업이익)));
      tr.appendChild(el('td', 'hide-sm ' + negCls(r.영업이익률), fPct(r.영업이익률)));
      const st = el('td', 'l hide-sm'); const strip = el('span', 'lx-strip');
      for (let j = 1; j <= 9; j++) { const v = r[j + '월']; const b = el('b', v ? '' : 'none', v === '등급외' ? '-' : (v || '·')); if (v) b.style.setProperty('--gc', GVAR[v]); b.title = `${j}월 ${v ? (v === '등급외' ? '등급외' : v + ' ' + GNAME[v]) : '거래 없음'}`; if (j > p.m) b.style.opacity = '.35'; strip.appendChild(b); }
      st.appendChild(strip); tr.appendChild(st);
      const rowOn = state.hlRows.has(r.고객코드);
      [...tr.children].forEach((td, i) => { const n = (rowOn ? 1 : 0) + (colOn(COLS[i]) ? 1 : 0); if (n) td.classList.add('hl' + n); });
      return tr;
    }));
    $('teamTableNote').textContent = `${state.top === '10' ? '매출 상위 ' + fInt(rows.length) + '곳 / 전체 ' + fInt(total) + '곳' : fInt(rows.length) + '곳 표시'} · 금액 단위: 백만원 · 숫자 항목 이름을 누르면 내림차순(▽) → 오름차순(△) 정렬과 열 음영, 거래처명을 누르면 행 음영 (겹치면 진하게) · 월별 등급은 각 월 당월 기준`;
  }

  function renderRule(p) {
    $('ruleTableTitle').textContent = '영업부서별 기준선 · ' + periodLabel(p);
    const body = $('bepBody');
    if (failed('bep')) { const tr = el('tr'); const td = el('td', 'lx-fail', 'Failed to load'); td.colSpan = 7; tr.appendChild(td); body.replaceChildren(tr); return; }
    const rows = dash.data('bep').data || [];
    const trs = [];
    for (const [bu, subs] of ORDER) for (const [, teams] of subs) for (const t of teams) {
      const r = rows.find((x) => x.key === `${p.mode}|${p.m}|${t}`); if (!r) continue;
      const tr = el('tr', 'click'); tr.dataset.key = r.key;
      tr.append(el('td', 'l', bu), el('td', 'l', t), el('td', null, fEok(r.매출)), el('td', negCls(r.공헌이익률), fPct(r.공헌이익률)), el('td', null, fPct(r.목표)), el('td', null, fPct(r.영업BEP)), el('td', null, fPct(r.직접BEP)));
      tr.addEventListener('click', () => goUnit('영업부서', t));
      trs.push(tr);
    }
    body.replaceChildren(...trs);
  }

  function renderExcl(p) {
    const body = $('exclBody');
    if (failed('excluded')) { const tr = el('tr'); const td = el('td', 'lx-fail', 'Failed to load'); td.colSpan = 8; tr.appendChild(td); body.replaceChildren(tr); return; }
    const rows = (dash.data('excluded').data || []).filter((x) => x.기준 === p.mode && x.월 === p.m).sort((a, b) => b.매출 - a.매출);
    body.replaceChildren(...rows.map((r) => {
      const tr = el('tr'); tr.dataset.key = r.key;
      tr.append(el('td', 'l', r.영업부서), el('td', null, fInt(r.거래처수)), el('td', negCls(r.매출), fEok(r.매출)), el('td', negCls(r.매출총이익), fEok(r.매출총이익)), el('td', negCls(r.공헌이익), fEok(r.공헌이익)), el('td', negCls(r.직접이익), fEok(r.직접이익)), el('td', negCls(r.영업이익), fEok(r.영업이익)), el('td', negCls(r.영업이익률), fPct(r.영업이익률)));
      return tr;
    }));
    if (!rows.length) { const tr = el('tr'); const td = el('td', 'lx-note', '이 기간에는 제외 영업부서 실적이 없습니다.'); td.colSpan = 8; tr.appendChild(td); body.replaceChildren(tr); }
  }

  function renderTake(p) {
    const take = $('take');
    const rows = dash.data('rollup').data || [];
    const f = (g) => rows.find((x) => x.기준 === p.mode && x.월 === p.m && x.단위 === '전사' && x.이름 === '전사' && x.등급 === g);
    const all = f('전체');
    if (!all) { take.textContent = '—'; return; }
    take.classList.remove('dash-skeleton');
    const parts = [`${periodLabel(p)} · 전체 ${fInt(all.거래처수)}곳`, ...GRADES.map((g) => { const r = f(g); return `${g}등급: ${fInt(r ? r.거래처수 : 0)}곳 (매출비중 ${fPct(r ? r.매출비중 : 0)})`; })];
    take.replaceChildren(...parts.flatMap((t, i) => { const sp = el('span', 'lx-tk', t); if (i > 0) { const g = GRADES[i - 1]; sp.style.setProperty('--gc', GVAR[g]); sp.classList.add('g'); } return i ? [el('span', 'lx-sep', '|'), sp] : [sp]; }));
  }


  /* ---------- main ---------- */
  function render() {
    const p = P();
    monthSel.value = String(p.m);
    $('modeM').setAttribute('aria-pressed', String(p.mode === 'M'));
    $('modeC').setAttribute('aria-pressed', String(p.mode === 'C'));
    $('basisSales').setAttribute('aria-pressed', String(state.basis === 'sales'));
    $('basisCount').setAttribute('aria-pressed', String(state.basis === 'count'));
    $('periodNote').textContent = 'RAW 26실적 1~9월 · ' + periodLabel(p);
    for (const t of ['summary', 'team', 'rule', 'excl']) { $(t).hidden = t !== p.tab; $('tab-' + t).setAttribute('aria-selected', String(t === p.tab)); }
    if (failed('pl')) { $('take').classList.remove('dash-skeleton'); $('take').textContent = '데이터를 불러오지 못했습니다.'; }
    renderTake(p);
    if (p.tab === 'summary') renderSummary(p);
    if (p.tab === 'team') renderTeam(p);
    if (p.tab === 'rule') renderRule(p);
    if (p.tab === 'excl') renderExcl(p);
  }

  let linked = false;
  dash.onData(() => {
    render();
    if (!linked && ready('rollup')) {
      linked = true;
      const h = location.hash.slice(1);
      if (['summary', 'team', 'rule', 'excl'].includes(h) && h !== P().tab) dash.setParams({ tab: h });
    }
  });
})();


(function () {
  var btn = document.getElementById('csvBtn');
  if (!btn) return;
  btn.addEventListener('click', function () {
    var rows = window.__custRows || [], info = window.__custInfo || {};
    var head = ['순위', '거래처', '고객코드', '영업부서', '등급', '매출(원)', '매출총이익(원)', '매출총이익률', '공헌이익(원)', '공헌이익률', '직접이익(원)', '직접이익률', '영업이익(원)', '영업이익률'];
    var keys = [null, '고객사명', '고객코드', '영업부서', '등급', '매출', '매출총이익', '매출총이익률', '공헌이익', '공헌이익률', '직접이익', '직접이익률', '영업이익', '영업이익률'];
    var cell = function (v) { if (v == null) return ''; var t = String(v); return /[",\r\n]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t; };
    var lines = [head.join(',')];
    rows.forEach(function (r, i) {
      lines.push(keys.map(function (k) {
        var v = k === null ? i + 1 : r[k];
        if (k && /률$/.test(k)) return v == null ? '' : (v * 100).toFixed(1) + '%';
        return cell(typeof v === 'number' ? Math.round(v) : v);
      }).join(','));
    });
    var name = 'customer_pl_' + (info.mode === 'M' ? 'month' + info.m : 'ytd1-' + info.m) + '.csv';
    var blob = new Blob(['\ufeff' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    var note = document.getElementById('teamTableNote');
    if (note) note.textContent = (info.team || '') + ' ' + (info.label || '') + ' 거래처 ' + rows.length + '곳을 ' + name + ' 파일로 저장했습니다 (다운로드 폴더).';
  });
})();

}
