// 처음 대시보드(Claude 대시보드 페이지)를 브라우저용 정적 파일로 변환한다.
//   full/dash.css  : 테마 토큰 + 글꼴(GitHub 상대경로) + 대시보드 스타일
//   full/dash.js   : dash API 대체(shim) + 화면 틀 삽입 + 대시보드 스크립트
// 데이터는 엑셀 매크로가 만든 화면 파일의 window.__CUSTPL_DATA 에서 읽는다.
const fs = require('fs');
const OUT = '/home/claude/customer-pl-addin/full/';
fs.mkdirSync(OUT + 'fonts', { recursive: true });
const src = fs.readFileSync('/home/claude/work/index.html', 'utf8');
const style = src.slice(src.indexOf('<style>') + 7, src.indexOf('</style>'));
const markup = src.slice(src.indexOf('</style>') + 8, src.indexOf('<script>')).trim();
const script = src.slice(src.indexOf('<script>') + 8, src.lastIndexOf('</script>'));

const FONT_DIR = '/home/claude/plex/plex/packages/plex-sans-kr/fonts/complete/woff2/hinted/';
const fonts = [['Regular', 400], ['Medium', 500], ['SemiBold', 600], ['Bold', 700]];
for (const [n] of fonts) fs.copyFileSync(FONT_DIR + `IBMPlexSansKR-${n}.woff2`, OUT + `fonts/IBMPlexSansKR-${n}.woff2`);
fs.copyFileSync(FONT_DIR + 'license.txt', OUT + 'fonts/OFL.txt');
fs.copyFileSync('/opt/npm-tools/node_modules/d3/dist/d3.min.js', OUT + 'd3.min.js');

const css = `/* 2026 거래처별 손익등급 - 브라우저용 (처음 대시보드와 같은 화면) */
${fonts.map(([n, w]) => `@font-face { font-family: 'IBM Plex Sans KR'; src: url('fonts/IBMPlexSansKR-${n}.woff2') format('woff2'); font-weight: ${w}; font-style: normal; font-display: swap; }`).join('\n')}
:root {
  --color-bg: #f6f7f9; --color-panel: #ffffff; --color-fg: #1b1f27; --color-fg-muted: #5f6673; --color-border-line: #dfe2e8;
  --color-ok: #1f7a3d; --color-warn: #a86400; --color-bad: #c0352b;
  --cds-chart-grid: #eef0f4; --cds-chart-muted: #9aa1ad; --cds-chart-axis: #c8ccd4; --cds-border: #dfe2e8; --cds-surface-popover: #ffffff; --cds-text-secondary: #5f6673;
  --cds-shadow-popover: 0 6px 20px rgba(20, 24, 33, .12);
  --cds-font-size-caption: 12px; --cds-font-size-body: 14px; --cds-font-size-heading: 16px; --cds-font-size-title: 22px; --cds-font-weight-medium: 500;
  --cds-gap-xs: 4px; --cds-gap-sm: 8px; --cds-gap-md: 12px; --cds-gap-lg: 20px;
  --cds-pad-xs: 4px; --cds-pad-sm: 8px; --cds-pad-md: 12px; --cds-pad-lg: 18px; --cds-radius: 8px; --cds-dur-slow: .3s;
  color-scheme: light;
}
@media (prefers-color-scheme: dark) { :root {
  --color-bg: #14171c; --color-panel: #1c2027; --color-fg: #e8eaee; --color-fg-muted: #a0a7b3; --color-border-line: #333944;
  --color-ok: #4cbb6c; --color-warn: #e0a03a; --color-bad: #f0695d;
  --cds-chart-grid: #262b34; --cds-chart-muted: #6b7380; --cds-chart-axis: #4a515d; --cds-border: #333944; --cds-surface-popover: #232831; --cds-text-secondary: #a0a7b3;
  --cds-shadow-popover: 0 6px 20px rgba(0, 0, 0, .45); color-scheme: dark; } }
[hidden] { display: none !important; }
html, body { margin: 0; background: var(--color-bg); color: var(--color-fg); font-family: 'IBM Plex Sans KR', 'Malgun Gothic', sans-serif; }
#dash-root { padding: 24px clamp(12px, 2.5vw, 32px); max-width: 1880px; margin: 0 auto; box-sizing: border-box; }
.dash-skeleton { color: transparent !important; background: var(--cds-chart-grid); border-radius: 4px; }
.full-msg { padding: 40px; font-size: 15px; color: var(--color-fg-muted); }
${style.replace(/@font-face[^\n]*\n/g, '')}
`;
fs.writeFileSync(OUT + 'dash.css', css);

const shim = `/* 2026 거래처별 손익등급 - 브라우저용 실행 코드 (처음 대시보드와 같은 화면)
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
  root.innerHTML = ${JSON.stringify(markup)};
  var t; window.addEventListener('resize', function () { clearTimeout(t); t = setTimeout(run, 150); });
})();
`;
fs.writeFileSync(OUT + 'dash.js', shim + '\nif (window.dash) {\n' + script + '\n}\n');
console.log('ok', fs.statSync(OUT + 'dash.css').size, fs.statSync(OUT + 'dash.js').size);
