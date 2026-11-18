'use strict';
/* 惠企政策智能体 · 运营看板 —— 第四轮：参考稿视觉（修正版） */

const $ = s => document.querySelector(s), $$ = s => [...document.querySelectorAll(s)];
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const NODES = ['进入','画像确认','匹配完成','看政策','选定政策','准备材料','办理指引','前往官方平台'];
const CODES = Object.fromEntries(
  ['画像放弃','查不到企业','无匹配结果','地区未覆盖','结果不吸引','生成失败','链接失效','低满意度']
    .map((v,i)=>['P'+String(i+1).padStart(2,'0'),v])
    .concat(['专利不足','软著不足','研发条件不足','缺少资质','材料写不出','不会办理','未注册企业','通过率过低','缺体系认证','经营异常']
      .map((v,i)=>['B'+String(i+1).padStart(2,'0'),v])));
const REQUIRED = '企业名称、所在地区、所属行业（国标代码）、企业规模、成立年限、企业类型、画像完整度、首次使用时间、最近使用时间、会话数、用过的功能、最远节点、卡点节点、卡点编码、主要缺口、推荐类目、推荐服务（服务名｜服务 ID）、长期培育候选、意向等级、触达授权、联系手机号、跟进人、跟进日期、跟进结果、成交服务、成交金额（元）'.split('、');
const FEATURES = [['企业画像','画像',['画像确认']],['政策匹配','匹配',['匹配完成']],['政策解读','解读',['查看政策详情']],['企业诊断','诊断',['诊断生成']],['申报材料','材料',['材料上传','材料智能生成']],['办理指引','指引',['进入办理指引']],['跳转官方平台','跳转',['确认跳转']]];
const DAYS = Array.from({length: 15}, (_, i) => `2026-11-${String(i + 1).padStart(2, '0')}`);
const reducedMotion = () => typeof matchMedia === 'function' && matchMedia('(prefers-reduced-motion: reduce)').matches;
/* 漏斗几何：列头 76px + 间隔 16px = 92；绘图区 298px（柱最大 280，顶部留 18px 给柱顶胶囊） */
const F_HEAD = 92, F_PLOT = 298, F_BAR = 280, F_BASE = 297;

/* ---------- 内联 SVG 图标（24 视框 / 1.5px 描边） ---------- */
const ICON_PATHS = {
  close: '<path d="M18 6 6 18"/><path d="m6 6 12 12"/>',
  'chevron-down': '<path d="m6 9 6 6 6-6"/>',
  'chevron-up': '<path d="m18 15-6-6-6 6"/>',
  sort: '<path d="m7 15 5 5 5-5"/><path d="m7 9 5-5 5 5"/>',
  external: '<path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/>'
};
const icon = (name, size = 14) =>
  `<svg class="icon" width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICON_PATHS[name]}</svg>`;

/* ---------- 状态 ---------- */
let data = null, filtered = [];
let filters = { intents: new Set() };
let detail = { codes: new Set(), node: '', cat: '', nurture: false };
let page = 1, sortKey = '', sortDir = 1, importId = 0, previousFocus = null;
let blockTab = 'B', funnelSel = -1;
let insightIdx = 0, insightTimer = null, insightSlides = [];
let periodShort = '—', metaFull = '', metaShort = '';

/* ---------- 小工具 ---------- */
const pct = (a, b) => b ? (a / b * 100).toFixed(1) + '%' : '—';
const name = c => c['企业名称'].replace(/^示例：/, '');
const cleanNode = v => String(v || '').replace(/^N\d+\s*/i, '').trim();
const list = v => String(v || '').split(/[、，,；]/).map(s => s.trim()).filter(Boolean);
const missing = s => `<p class="notice">导入文件里没有「${s}」表，此处无数据</p>`;
const fmtShort = t => String(t || '').slice(5, 16);
const codeChip = k => `<span class="code-chip ${k[0] === 'B' ? 'b' : 'p'}">${esc(k)}</span>`;
const codeName = k => esc(CODES[k] || '未知卡点');
const downloadCsv = (title, rows) => {
  const csv = '\uFEFF' + rows.map(r => r.map(v => `"${String(v ?? '').replace(/"/g, '""')}"`).join(',')).join('\r\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
  a.download = title + '.csv';
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
};

/* ---------- 导入与解析 ---------- */
function readSheet(w, s) {
  if (!w.Sheets[s]) return null;
  const rows = XLSX.utils.sheet_to_json(w.Sheets[s], { header: 1, defval: '', raw: false });
  const at = rows.findIndex(r => r.includes('企业名称') || r.includes('事件时间'));
  if (at < 0) return { headers: [], rows: [] };
  const headers = rows[at].map(v => String(v).trim());
  return { headers, rows: rows.slice(at + 1).filter(r => r.some(v => String(v).trim())).map(r => Object.fromEntries(headers.map((h, i) => [h, String(r[i] ?? '').trim()]))) };
}
function parse(buffer) {
  const signature = new Uint8Array(buffer);
  if (signature[0] !== 80 || signature[1] !== 75) throw Error('invalid xlsx');
  const w = XLSX.read(buffer, { type: 'array' }), c = readSheet(w, '企业明细');
  if (!c) throw Error('导入文件里没有「企业明细」表，此处无数据');
  const absent = REQUIRED.filter(h => !c.headers.includes(h));
  if (absent.length) throw Error('企业明细缺少以下列：' + absent.join('、') + '，请用最新导入模版导出');
  const events = readSheet(w, '事件明细'), gaps = readSheet(w, '条件缺口明细');
  const sumStr = w.Sheets['汇总'] ? XLSX.utils.sheet_to_json(w.Sheets['汇总'], { header: 1, defval: '' })[0]?.join(' ') : '';
  const range = sumStr?.match(/统计时间段[：:]\s*(\S+)\s+至\s+(\S+)/);
  const etimes = events?.rows.map(e => e['事件时间']).filter(Boolean).sort() || [];
  return {
    companies: c.rows.map(c => ({ ...c, level: NODES.indexOf(cleanNode(c['最远节点'])) + 1, codes: list(c['卡点编码']), cats: list(c['推荐类目']) })),
    events: events?.rows.map(e => { let fields = null; try { fields = JSON.parse(e['关键字段']); if (!fields || typeof fields !== 'object' || Array.isArray(fields)) fields = null; } catch {} return { ...e, fields }; }) || [],
    gaps: gaps?.rows || [], hasEvents: !!events, hasGaps: !!gaps, hasSummary: !!w.Sheets['汇总'],
    summary: sumStr,
    periodStart: range ? range[1] : etimes[0]?.slice(0, 10) || '', periodEnd: range ? range[2] : etimes.at(-1)?.slice(0, 10) || ''
  };
}
function fitMeta() {
  const el = $('#meta');
  if (!metaFull) return;
  el.textContent = metaFull;
  el.title = metaFull;
  if (el.scrollWidth > el.clientWidth && metaShort) el.textContent = metaShort;
}
async function loadFile(file) {
  const id = ++importId;
  $('#error').hidden = true;
  try {
    if (!/\.xlsx$/i.test(file.name)) throw Error();
    const next = parse(await file.arrayBuffer());
    if (id !== importId) return;
    data = next;
    reset(false);
    $('#region').innerHTML = '<option value="">全部地区</option>' +
      [...new Set(data.companies.map(c => c['所在地区']))].filter(Boolean).sort((a, b) => a.localeCompare(b, 'zh-CN')).map(r => `<option>${esc(r)}</option>`).join('');
    const range = data.summary?.match(/统计时间段[：:]\s*(\S+)\s+至\s+(\S+)/);
    const times = data.events.map(e => e['事件时间']).filter(Boolean).sort();
    const a = range ? range[1] : times[0]?.slice(0, 10), b = range ? range[2] : times.at(-1)?.slice(0, 10);
    periodShort = a && b ? a.slice(5) + ' – ' + b.slice(5) : '—';
    $('#period-text').textContent = periodShort;
    metaFull = `数据文件：${file.name} · 统计时间段：${a ?? '—'} 至 ${b ?? '—'} · 企业 ${data.companies.length} 家 · 事件 ${data.events.length.toLocaleString()} 条`;
    metaShort = `${file.name} · ${periodShort} · ${data.companies.length} 家 · ${data.events.length.toLocaleString()} 条事件`;
    fitMeta();
    $('#empty').hidden = true; $('#dashboard').hidden = false; $('#head-controls').hidden = false;
    document.body.classList.remove('empty-mode');
    render();
    showToast(`已导入 ${file.name}，共 ${data.companies.length} 家企业`);
  } catch (e) {
    if (id !== importId) return;
    data = null;
    $('#dashboard').hidden = true; $('#head-controls').hidden = true; $('#empty').hidden = false; $('#meta').textContent = ''; metaFull = metaShort = '';
    document.body.classList.add('empty-mode');
    $('#error').hidden = false;
    $('#error').textContent = /^企业明细缺少|^导入文件里没有/.test(e.message) ? e.message : '文件无法读取，请确认是后台导出的 .xlsx 文件';
  } finally {
    $('#file').value = ''; closeDrawer();
  }
}
function reset(draw = true) {
  filters = { intents: new Set() };
  detail = { codes: new Set(), node: '', cat: '', nurture: false };
  page = 1; sortKey = ''; sortDir = 1; funnelSel = -1; insightIdx = 0;
  $('#region').value = ''; $('#auth').value = ''; $('#search').value = '';
  $$('[data-intent]').forEach(b => b.setAttribute('aria-pressed', 'false'));
  $$('[data-code]').forEach(c => c.checked = false);
  syncCodeBtn();
  if (draw && data) render();
}
addEventListener('resize', () => { fitMeta(); showFunnelTip($(`.funnel-col[data-col="${funnelSel}"]`)); });

/* ---------- 渲染总入口 ---------- */
function render() {
  filtered = data.companies.filter(c =>
    (!$('#region').value || c['所在地区'] === $('#region').value) &&
    (!filters.intents.size || filters.intents.has(c['意向等级'])) &&
    (!$('#auth').value || c['触达授权'] === $('#auth').value));
  renderFunnel(); renderLeads(); renderDaily(); renderEvents(); renderOppDots(); renderInsight();
  renderBlocks(); renderNurture(); renderOpps(); renderFeatures(); renderProcess(); renderDetail();
}

/* ---------- 漏斗统计（多处复用） ---------- */
function funnelStats() {
  const n = filtered.length;
  const counts = NODES.map((_, i) => filtered.filter(c => i === 5 ? c['用过的功能'].includes('材料') : c.level >= i + 1).length);
  const stuck = NODES.map(v => filtered.filter(c => cleanNode(c['卡点节点']) === v).length);
  const conv = i => i ? pct(counts[i], counts[i === 6 ? 4 : i - 1]) : '—';
  return { n, counts, stuck, conv };
}
function renderFunnel() {
  const { n, counts, stuck, conv } = funnelStats();
  if (funnelSel < 0 || funnelSel > 7) funnelSel = stuck.indexOf(Math.max(...stuck));
  if (!n) { $('#funnel').innerHTML = '<p class="no-data">暂无符合条件的企业</p>'; return; }
  const ticks = [0, 25, 50, 75, 100].map(p => `<span style="bottom:${Math.round(1 + F_BAR * p / 100)}px">${p === 0 ? '0' : p + '%'}</span>`).join('');
  const grid = ''; /* 不画网格线，保持斜坡缝隙连续 */
  const h = i => Math.round(counts[i] / n * F_BAR);
  const cols = NODES.map((v, i) =>
    `<button type="button" class="funnel-col${i === funnelSel ? ' sel' : ''}" data-col="${i}" aria-label="${v}，到达 ${counts[i]} 家">
      <span class="fc-head"><span class="fc-name">${v}</span><span class="fc-num">${counts[i]}</span></span>
      <span class="fc-plot"><i class="fc-bar" style="height:${h(i)}px"></i>${counts[i] ? `<i class="fc-mark" style="bottom:${h(i) + 11}px"></i>` : ''}</span>
    </button>`).join('');
  const polys = NODES.slice(0, -1).map((_, i) => {
    const x1 = i * 100 + 86, x2 = (i + 1) * 100;
    /* 注意：此 WebKit 不渲染 attribute 里的半透明填色，必须用 style 填色 + fill-opacity 分离写法 */
    return `<polygon points="${x1},${F_BASE - h(i)} ${x2},${F_BASE - h(i + 1)} ${x2},${F_BASE} ${x1},${F_BASE}" style="fill: var(--c-neg)" fill-opacity=".25"/>` +
      `<path d="M ${x1} ${F_BASE - h(i)} L ${x2} ${F_BASE - h(i + 1)}" style="stroke: var(--c-neg)" stroke-opacity=".5" stroke-width="1" fill="none" vector-effect="non-scaling-stroke"/>`;
  }).join('');
  const i = funnelSel, v = NODES[i];
  const band = `<div class="funnel-band">
    <span class="fb-text"><b>${v}</b> 有 <b>${stuck[i]}</b> 家企业停在这里，转化率 ${conv(i)}</span>
    <button type="button" class="pill-dark" data-band-stuck="${v}" ${stuck[i] ? '' : 'disabled'}>查看这 ${stuck[i]} 家</button>
  </div>`;
  $('#funnel').innerHTML = (data.hasSummary ? '' : missing('汇总')) +
    `<div class="funnel-wrap">
      <div class="funnel-y" aria-hidden="true">${ticks}</div>
      <div class="funnel-scroll" id="funnel-scroll">
        <div class="funnel-cols">${grid}<svg class="slopes" viewBox="0 0 800 ${F_PLOT}" preserveAspectRatio="none" aria-hidden="true">${polys}</svg>${cols}</div>
        <div class="ftip" id="ftip"></div>
      </div>
    </div>${band}`;
  showFunnelTip($(`.funnel-col[data-col="${funnelSel}"]`)); /* 选中列常显胶囊提示 */
}
function showFunnelTip(col) {
  const tip = $('#ftip'), scroll = $('#funnel-scroll');
  if (!col || !tip || !scroll) return;
  const { counts, stuck, conv } = funnelStats();
  const i = Number(col.dataset.col);
  tip.innerHTML = `到达 <b>${counts[i]}</b> 家<span class="sep">｜</span>转化率 <b>${conv(i)}</b><span class="sep">｜</span>卡住 <b class="neg">${stuck[i]}</b> 家`;
  const n = filtered.length || 1;
  const sr = scroll.getBoundingClientRect(), cr = col.getBoundingClientRect();
  const scrollL = scroll.scrollLeft;
  const colL = cr.left - sr.left + scrollL, colR = cr.right - sr.left + scrollL; /* 内容坐标，横滚后仍正确 */
  const tw = tip.offsetWidth, th = tip.offsetHeight;
  const h = Math.round(counts[i] / n * F_BAR);
  const barTop = F_HEAD + (F_BASE - h); /* 选中柱顶在滚动容器内的 y */
  /* 按实际剩余空间判断：选中列右缘到绘图区可视右缘 ≥ 胶囊宽 + 8 → 右侧；否则左侧，并抬高到所覆盖柱顶之上（不压柱） */
  const fitsRight = colR + 8 + tw <= scrollL + sr.width - 2;
  let left, top;
  if (fitsRight) {
    left = colR + 8;
    top = Math.max(2, Math.min(barTop - th / 2, sr.height - th - 2));
  } else {
    left = Math.max(2, colL - 8 - tw);
    const colW = scroll.scrollWidth / 8;
    const i0 = Math.max(0, Math.floor(left / colW)), i1 = Math.min(7, Math.ceil((left + tw) / colW) - 1);
    let minTop = Infinity;
    for (let k = i0; k <= i1; k++) minTop = Math.min(minTop, F_HEAD + (F_BASE - Math.round(counts[k] / n * F_BAR)));
    top = Math.max(2, Math.min(minTop - th - 6, barTop - Math.round(th / 2)));
  }
  tip.style.left = left + 'px';
  tip.style.top = top + 'px';
  tip.classList.add('show');
}
const hideFunnelTipToSel = () => showFunnelTip($(`.funnel-col[data-col="${funnelSel}"]`));

/* ---------- 线索有效率 ---------- */
function renderLeads() {
  const n = filtered.length;
  const profile = filtered.filter(c => c.level >= 2).length;
  const leads = filtered.filter(c => ['高','中'].includes(c['意向等级'])).length;
  const deals = filtered.filter(c => c['成交服务']).length;
  const rows = [['高意向','高','var(--stripe-pos-strong)'], ['中意向','中','var(--stripe-pos)'], ['低意向','低','var(--stripe-neutral)']]
    .map(([label, key, grad]) => {
      const count = filtered.filter(c => c['意向等级'] === key).length;
      return `<div class="lead-row"><div class="lr-top"><span class="lr-label">${label}</span><span class="lr-val">${count} 家</span></div>
        <div class="lr-track"><div class="lr-fill" style="width:${profile ? count / profile * 100 : 0}%;background:${grad}"></div></div></div>`;
    }).join('');
  $('#leads').innerHTML =
    `<div class="leads-big"><span class="big big72">${pct(leads, profile)}</span><span class="north">北极星</span></div>
     <div class="leads-div"></div>${rows}
     <div class="lead-foot">服务转化率 ${deals ? pct(deals, leads) : '—'} · 业务回填</div>
     <p class="lead-rule">高：点过服务、找过人工，或选了政策且缺条件；中：看过 2 条以上政策，或进了办理指引没去官方平台；14 天无操作降一级</p>`;
}

/* ---------- 每日进入企业 ---------- */
function dailyNewCounts() {
  const counts = Object.fromEntries(DAYS.map(d => [d, 0]));
  const inPeriod = d => !data.periodStart || (d >= data.periodStart && d <= data.periodEnd);
  if (data.hasEvents) {
    const names = new Set(filtered.map(c => c['企业名称']));
    const first = new Map();
    for (const e of data.events) {
      const n = e['企业名称'], d = (e['事件时间'] || '').slice(0, 10);
      if (names.has(n) && d && (!first.has(n) || d < first.get(n))) first.set(n, d);
    }
    for (const d of first.values()) if (d in counts && inPeriod(d)) counts[d]++;
  } else {
    for (const c of filtered) { const d = String(c['首次使用时间'] || '').slice(0, 10); if (d in counts && inPeriod(d)) counts[d]++; }
  }
  return DAYS.map(d => counts[d]);
}
function renderDaily() {
  const counts = dailyNewCounts();
  const total = counts.reduce((a, b) => a + b, 0);
  const el = $('#daily');
  if (!total) { el.innerHTML = '<p class="no-data">暂无进入记录</p>'; return; }
  const W = 300, H = 150, top = 12, bottom = 144;
  const max = Math.max(...counts);
  const x = i => 10 + i * 280 / 14, y = v => bottom - v / max * (bottom - top);
  let line = `M ${x(0)} ${y(counts[0])}`;
  for (let i = 1; i < 15; i++) line += ` L ${x(i)} ${y(counts[i - 1])} L ${x(i)} ${y(counts[i])}`;
  const area = line + ` L ${x(14)} ${bottom} L ${x(0)} ${bottom} Z`;
  const peak = counts.indexOf(max);
  const capLeft = Math.min(88, Math.max(12, x(peak) / W * 100));
  const labels = [0, 3, 7, 10, 14].map((i, k) => k === 4
    ? `<span style="left:auto;right:0;transform:none">${DAYS[i].slice(5)}</span>`
    : `<span style="left:${(x(i) / W * 100).toFixed(1)}%">${DAYS[i].slice(5)}</span>`).join('');
  el.innerHTML =
    `<div class="chart-body">
      <div class="chart">
        <svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="none" aria-hidden="true">
          <defs><pattern id="dstripe" width="4" height="4" patternUnits="userSpaceOnUse"><rect width="1" height="4" style="fill: var(--c-pos)" opacity=".15"/></pattern></defs>
          <path d="${area}" fill="url(#dstripe)"/>
          <path d="${line}" fill="none" style="stroke: var(--c-pos)" stroke-width="2" vector-effect="non-scaling-stroke" stroke-linejoin="round"/>
          <circle cx="${x(peak)}" cy="${y(max)}" r="4" style="fill: var(--c-pos-strong)" vector-effect="non-scaling-stroke"/>
        </svg>
        <span class="chart-cap" style="left:${capLeft.toFixed(1)}%;top:${(y(max) / H * 100 - 4).toFixed(1)}%">${max} 家</span>
      </div>
      <div class="x-axis">${labels}</div>
    </div>`;
}

/* ---------- 事件量（点状柱） ---------- */
function renderEvents() {
  const el = $('#events');
  if (!data.hasEvents) { el.innerHTML = '<p class="no-data">暂无事件数据</p>'; return; }
  const names = new Set(filtered.map(c => c['企业名称']));
  const events = data.events.filter(e => names.has(e['企业名称']));
  const counts = DAYS.map(d => events.filter(e => (e['事件时间'] || '').slice(0, 10) === d).length);
  const total = events.length;
  const peak = counts.indexOf(Math.max(...counts));
  const cols = counts.map((c, i) => {
    const dots = Array.from({length: Math.min(6, Math.ceil(c / 30))}, () => '').map(() => `<span class="dot${i === peak ? ' on' : ''}"></span>`).join('');
    const cap = i === peak ? `<span class="dot-cap">峰值：${DAYS[i].slice(5)}</span>` : '';
    return `<div class="dot-col" data-tip="${DAYS[i].slice(5)} · ${c} 条">${cap}${dots}</div>`;
  }).join('');
  el.innerHTML =
    `<div class="dot-flex">
      <div class="dot-left"><span class="big big64">${total.toLocaleString()}</span><small>条事件</small></div>
      <div class="dot-grid">${cols}</div>
      <div class="dot-right"><small>日均</small><span class="num22">${(total / 15).toFixed(1)}</span></div>
    </div>`;
}

/* ---------- 商机企业（点状柱） ---------- */
function renderOppDots() {
  const cats = [...new Set(filtered.flatMap(c => c.cats))]
    .map(k => [k, filtered.filter(c => c.cats.includes(k))]).sort((a, b) => b[1].length - a[1].length);
  const biz = filtered.filter(c => c.codes.some(k => k[0] === 'B'));
  const el = $('#opp-dots');
  if (!biz.length) { el.innerHTML = '<p class="no-data">暂无商机企业</p>'; return; }
  const top = cats[0];
  const cols = cats.map(([k, cs]) =>
    `<div class="dot-col${cs === top[1] ? ' on' : ''}" data-tip="${esc(k)} · ${cs.length} 家">
      ${cs === top[1] ? `<span class="dot-cap">最多：${esc(k)}</span>` : ''}
      ${Array.from({length: Math.min(8, cs.length)}, () => `<span class="dot${cs === top[1] ? ' on' : ''}"></span>`).join('')}
    </div>`).join('');
  el.innerHTML =
    `<div class="dot-flex">
      <div class="dot-left"><span class="big big64">${biz.length}</span><small>家商机企业</small></div>
      <div class="dot-grid">${cols}</div>
      <div class="dot-right"><small>已授权</small><span class="num22">${biz.filter(c => c['触达授权'] === '已授权').length}</span></div>
    </div>`;
}

/* ---------- 洞察卡 ---------- */
function renderInsight() {
  const { counts, stuck } = funnelStats();
  const profile = counts[1], biz = filtered.filter(c => c.codes.some(k => k[0] === 'B'));
  const ns = filtered.filter(c => c['长期培育候选'] === '是');
  let worst = 0, worstDrop = -1;
  for (let i = 1; i < 8; i++) {
    const denom = counts[i === 6 ? 4 : i - 1];
    if (!denom) continue;
    const drop = 1 - counts[i] / denom;
    if (drop > worstDrop) { worstDrop = drop; worst = i; }
  }
  const prev = NODES[worst === 6 ? 4 : worst - 1] || NODES[0];
  const a = ns.filter(c => c['触达授权'] === '已授权').length;
  insightSlides = [
    worstDrop >= 0
      ? { num: (worstDrop * 100).toFixed(1) + '%', title: `${prev}到${NODES[worst]}流失最多`, desc: `这一步有 ${stuck[worst]} 家企业停下，先看他们缺什么材料。` }
      : { num: '—', title: '暂无转化数据', desc: '导入数据后这里会给出流失最多的步骤。' },
    { num: pct(biz.length, profile), title: `完成画像的企业里，${pct(biz.length, profile)} 有可跟进的商机`, desc: `其中已授权 ${biz.filter(c => c['触达授权'] === '已授权').length} 家，可以直接联系。` },
    { num: `${ns.length}<span class="unit">家</span>`, title: '适合做多年培育规划', desc: `已授权 ${a} 家，未授权 ${ns.length - a} 家。` }
  ];
  if (insightIdx >= insightSlides.length) insightIdx = 0;
  paintInsight();
  $('#insight-segs').innerHTML = insightSlides.map((_, i) =>
    `<button type="button" class="${i === insightIdx ? 'on' : ''}" data-insight="${i}" aria-label="第 ${i + 1} 条洞察"></button>`).join('');
  clearInterval(insightTimer);
  if (!reducedMotion() && insightSlides.length > 1) insightTimer = setInterval(() => showInsight((insightIdx + 1) % insightSlides.length), 6000);
}
function paintInsight() {
  const s = insightSlides[insightIdx];
  $('#insight-slide').innerHTML =
    `<div class="insight-num">${s.num}</div><div class="insight-title">${s.title}</div><div class="insight-desc">${s.desc}</div>`;
  $$('#insight-segs button').forEach((b, i) => b.classList.toggle('on', i === insightIdx));
}
function showInsight(i) {
  if (i === insightIdx) { paintInsight(); return; }
  insightIdx = i;
  if (reducedMotion()) { paintInsight(); return; }
  const slide = $('#insight-slide');
  slide.classList.add('fade');
  setTimeout(() => { paintInsight(); slide.classList.remove('fade'); }, 300);
}

/* ---------- 卡点分布（胶囊 tab + 两列列表） ---------- */
function renderBlocks() {
  const n = filtered.length;
  const sideCount = p => filtered.filter(c => c.codes.some(k => k[0] === p)).length;
  const build = p => {
    const counted = Object.keys(CODES).filter(k => k[0] === p)
      .map(k => [k, filtered.filter(c => c.codes.includes(k)).length])
      .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
    const max = Math.max(...counted.map(x => x[1]), 1);
    const row = ([k, cnt]) =>
      `<button type="button" class="block-row${p === 'P' ? ' p' : ''}${cnt ? '' : ' zero'}" data-block="${k}">
        <span class="b-code">${k}</span><span class="b-name">${CODES[k]}</span>
        <span class="track"><i style="width:${cnt / max * 100}%"></i></span>
        <span class="b-count">${cnt}</span><span class="b-pct">${pct(cnt, n)}</span>
      </button>`;
    const half = Math.ceil(counted.length / 2);
    return `<div class="block-col">${counted.slice(0, half).map(row).join('')}</div><div class="block-col">${counted.slice(half).map(row).join('')}</div>`;
  };
  $('#blocks').innerHTML =
    `<div class="block-tabs" role="tablist" aria-label="卡点分布">
      <button type="button" role="tab" aria-pressed="${blockTab === 'B'}" data-tab="B">企业侧 · 商机 ${sideCount('B')}</button>
      <button type="button" role="tab" aria-pressed="${blockTab === 'P'}" data-tab="P">产品侧 · 产品优化 ${sideCount('P')}</button>
    </div><div class="block-cols">${build(blockTab)}</div>`;
}

/* ---------- 长期培育候选 ---------- */
function renderNurture() {
  const ns = filtered.filter(c => c['长期培育候选'] === '是');
  const a = ns.filter(c => c['触达授权'] === '已授权').length, b = ns.length - a;
  const bar = (label, val, grad) =>
    `<div class="nr-row"><div class="nr-top"><span>${label}</span><b>${val} 家</b></div>
     <div class="nr-track"><i style="width:${ns.length ? val / ns.length * 100 : 0}%;background:${grad}"></i></div></div>`;
  $('#nurture').innerHTML =
    `<div class="nurture-big"><span class="big big72">${ns.length}</span><span class="unit">家</span></div>
     ${ns.length ? '<div class="nurture-rows">' + bar('已授权', a, 'var(--stripe-pos)') + bar('未授权', b, 'var(--stripe-neutral)') + '</div>' : ''}
     <button type="button" class="pill-dark" id="view-nurture">查看名单</button>`;
}

/* ---------- 商机分布 ---------- */
function renderOpps() {
  const cats = [...new Set(filtered.flatMap(c => c.cats))]
    .map(k => [k, filtered.filter(c => c.cats.includes(k))]).sort((a, b) => b[1].length - a[1].length);
  const max = Math.max(...cats.map(x => x[1].length), 1);
  $('#opportunities').innerHTML = cats.length
    ? `<table><thead><tr><th>推荐类目</th><th class="num">候选企业数</th><th class="num">其中已授权</th><th class="num">其中已成交</th><th></th></tr></thead><tbody>` +
      cats.map(([k, cs]) => {
        const auth = cs.filter(c => c['触达授权'] === '已授权').length;
        const deal = cs.filter(c => c['成交服务']).length;
        return `<tr><td>${esc(k)}</td>
          <td class="num"><div class="opp-num"><span>${cs.length}</span><span class="opp-bar"><i style="width:${cs.length / max * 100}%"></i></span></div></td>
          <td class="num">${auth}</td>
          <td class="num">${deal ? `<span class="success">${deal}</span>` : '0'}</td>
          <td class="num"><button type="button" class="cap xs" data-cat="${esc(k)}">看企业</button></td></tr>`;
      }).join('') + `</tbody></table>`
    : '<p class="no-data">暂无推荐类目</p>';
}

/* ---------- 功能使用 / 过程指标 ---------- */
function renderFeatures() {
  const n = filtered.length;
  const names = new Set(filtered.map(c => c['企业名称']));
  const events = data.events.filter(e => names.has(e['企业名称']));
  $('#features').innerHTML = (!data.hasEvents ? missing('事件明细') : '') +
    FEATURES.map(([label, key, ev]) => {
      const count = filtered.filter(c => c['用过的功能'].includes(key)).length;
      const times = data.hasEvents ? events.filter(e => ev.includes(e['事件'])).length : null;
      return `<div class="feat-row"><span class="feat-name">${label}</span>
        <span class="feat-track"><i style="width:${n ? count / n * 100 : 0}%"></i></span>
        <span class="feat-val"><b>${pct(count, n)}</b> · ${times === null ? '—' : times + ' 次'}</span></div>`;
    }).join('');
}
function processMetrics() {
  const names = new Set(filtered.map(c => c['企业名称']));
  const valid = data.events.filter(e => names.has(e['企业名称']) && e.fields);
  const get = ev => valid.filter(e => e['事件'] === ev);
  const f = (e, cn, en) => e.fields[cn] ?? e.fields[en];
  const is = (e, cn, en, values) => values.includes(f(e, cn, en));
  const match = get('匹配完成'), relaxed = match.filter(e => f(e, '是否放宽', 'relaxed') === true);
  const mats = get('材料智能生成'), shown = get('服务入口展示');
  const stars = get('满意度评价').filter(e => Number.isFinite(Number(f(e, '星级', 'stars'))));
  const unique = es => new Set(es.map(e => e['企业名称'])).size;
  return [
    ['无匹配率', match.filter(e => Number(f(e, '结果数', 'result_count')) === 0).length, match.length, '结果数为 0 ÷ 匹配完成'],
    ['放宽后有结果率', relaxed.filter(e => f(e, '结果数', 'result_count') > 0).length, relaxed.length, '放宽后有结果 ÷ 放宽匹配'],
    ['平均匹配政策数', match.reduce((s, e) => s + (Number(f(e, '结果数', 'result_count')) || 0), 0), match.length, '匹配结果数平均值', 'avg'],
    ['匹配结果点击率', get('查看政策详情').length, get('政策列表曝光').length, '查看政策详情 ÷ 政策列表曝光'],
    ['材料生成失败率', mats.filter(e => !is(e, '结果', 'result', ['成功', 'success'])).length, mats.length, '生成失败 ÷ 材料智能生成'],
    ['跳转成功率', get('跳转结果').filter(e => is(e, '结果', 'result', ['已打开', 'opened'])).length, get('确认跳转').filter(e => is(e, '操作', 'action', ['确认', 'confirm'])).length, '已打开 ÷ 确认跳转'],
    ['服务入口点击率', get('服务入口点击').length, shown.length, '点击 ÷ 服务入口展示'],
    ['服务触达打扰率', unique(valid.filter(e => e['事件'] === '暂不需要' || e['事件'] === '触达授权变更' && is(e, '操作', 'action', ['关闭', 'revoke']))), unique(shown), '拒绝或关闭企业 ÷ 展示企业'],
    ['用户满意度', stars.reduce((s, e) => s + Number(f(e, '星级', 'stars')), 0), stars.length, `共 ${stars.length} 条评价`, 'stars']
  ];
}
function renderProcess() {
  if (!data.hasEvents) { $('#process').innerHTML = '<p class="no-data">暂无事件数据</p>'; return; }
  /* 明显偏差项用负向文字色：无匹配率 >10%、材料生成失败率 >10%、打扰率 >20%、跳转成功率 <90% */
  const NEG = new Set(['无匹配率', '材料生成失败率', '服务触达打扰率', '跳转成功率']);
  const isBad = (label, ratio) => {
    if (!NEG.has(label) || !ratio) return false;
    return label === '跳转成功率' ? ratio < 0.9 : ratio > 0.1 && (label !== '服务触达打扰率' || ratio > 0.2);
  };
  $('#process').innerHTML = processMetrics().map(([label, a, b, desc, type]) => {
    const val = b ? (type ? (a / b).toFixed(1) + (type === 'stars' ? ' 星' : '') : pct(a, b)) : '—';
    const neg = b && isBad(label, a / b);
    return `<div><h4>${label}</h4><strong class="${neg ? 'neg' : ''}">${val}</strong><small>${b ? desc : '暂无相关事件'}</small></div>`;
  }).join('');
}

/* ---------- 企业明细 ---------- */
function renderDetail() {
  const q = $('#search').value.trim();
  let cs = filtered.filter(c =>
    (!detail.node || cleanNode(c['卡点节点']) === detail.node) &&
    ([...detail.codes].every(k => c.codes.includes(k))) &&
    (!detail.cat || c.cats.includes(detail.cat)) &&
    (!detail.nurture || c['长期培育候选'] === '是') &&
    name(c).includes(q));
  const rank = { 高: 0, 中: 1, 低: 2 };
  cs.sort((a, b) => {
    if (sortKey) {
      const key = k => sortKey === '最远节点' ? k.level : sortKey === '意向等级' ? (rank[k['意向等级']] ?? 3) : sortKey === '企业名称' ? name(k) : k[sortKey];
      const va = key(a), vb = key(b);
      return (typeof va === 'number' ? va - vb : String(va).localeCompare(String(vb), 'zh-CN')) * sortDir;
    }
    return (rank[a['意向等级']] ?? 3) - (rank[b['意向等级']] ?? 3) || b['最近使用时间'].localeCompare(a['最近使用时间']);
  });
  const pages = Math.max(1, Math.ceil(cs.length / 20));
  page = Math.min(page, pages);
  $('#count').textContent = `共 ${cs.length} 家`;
  renderChips();
  $$('[data-code]').forEach(el => el.checked = detail.codes.has(el.dataset.code));
  if (!cs.length) {
    $('#company-table').innerHTML = '<div class="no-results">没有符合条件的企业<button type="button" class="pill-dark" id="clear-all">清除筛选</button></div>';
    $('#pagination').innerHTML = '';
    return;
  }
  const sortIcon = key => icon(sortKey === key ? (sortDir === 1 ? 'chevron-up' : 'chevron-down') : 'sort', 12);
  const headers = ['企业名称','最远节点','卡点编码','主要缺口','推荐类目','意向等级','触达授权','长期培育','最近使用'];
  const sortKeys = { '企业名称':'企业名称', '最远节点':'最远节点', '意向等级':'意向等级', '最近使用':'最近使用时间' };
  $('#company-table').innerHTML = '<table><thead><tr>' +
    headers.map(h => `<th>${sortKeys[h] ? `<button type="button" data-sort="${sortKeys[h]}">${h} ${sortIcon(sortKeys[h])}</button>` : h}</th>`).join('') +
    '</tr></thead><tbody>' +
    cs.slice((page - 1) * 20, page * 20).map(c => {
      const sub = [c['所在地区'], c['所属行业（国标代码）']].filter(Boolean).join(' · ');
      const intent = c['意向等级'];
      const dot = `<span class="dot ${intent === '高' ? 'hi' : intent === '中' ? 'md' : 'lo'}"></span>${esc(intent)}`;
      return `<tr tabindex="0" data-company="${data.companies.indexOf(c)}" aria-label="${esc(name(c))}">` +
        `<td class="c-name"><div class="nm" title="${esc(name(c))}">${esc(name(c))}</div><div class="sub" title="${esc(sub)}">${esc(sub) || '—'}</div></td>` +
        `<td>${esc(cleanNode(c['最远节点']))}</td>` +
        `<td><div class="c-codes">${c.codes.map(codeChip).join('')}</div></td>` +
        `<td class="ellipsis" title="${esc(c['主要缺口'])}">${esc(c['主要缺口']) || '—'}</td>` +
        `<td class="ellipsis" title="${esc(c['推荐类目'])}">${esc(c['推荐类目']) || '—'}</td>` +
        `<td>${dot}</td>` +
        `<td class="${c['触达授权'] === '已授权' ? 'success' : 'muted'}">${esc(c['触达授权'])}</td>` +
        `<td>${c['长期培育候选'] === '是' ? '<span class="nurture-chip">培育候选</span>' : ''}</td>` +
        `<td title="${esc(c['最近使用时间'])}">${esc(fmtShort(c['最近使用时间']))}</td></tr>`;
    }).join('') + '</tbody></table>';
  $('#pagination').innerHTML =
    `<button type="button" class="cap xs" data-page="-1" ${page === 1 ? 'disabled' : ''}>上一页</button>` +
    `<span>第 ${page} / ${pages} 页</span>` +
    `<button type="button" class="cap xs" data-page="1" ${page === pages ? 'disabled' : ''}>下一页</button>`;
}
function renderChips() {
  const chips = [];
  if ($('#region').value) chips.push(['region', '地区：' + $('#region').value]);
  [...filters.intents].sort().forEach(k => chips.push(['intent:' + k, '意向等级：' + k]));
  if ($('#auth').value) chips.push(['auth', '触达授权：' + $('#auth').value]);
  if (detail.node) chips.push(['node', '卡点节点：' + detail.node]);
  [...detail.codes].sort().forEach(k => chips.push(['code:' + k, '卡点编码：' + k]));
  if (detail.cat) chips.push(['cat', '推荐类目：' + detail.cat]);
  if (detail.nurture) chips.push(['nurture', '长期培育候选']);
  const q = $('#search').value.trim();
  if (q) chips.push(['search', '搜索：' + q]);
  $('#chips').innerHTML = chips.map(([k, l]) =>
    `<button type="button" class="chip" data-clear="${esc(k)}">${esc(l)}${icon('close', 12)}</button>`).join('') +
    (chips.length ? `<button type="button" class="text-btn" id="clear-all">清除全部</button>` : '');
}

/* ---------- 卡点编码浮层 ---------- */
function buildPopover() {
  $('#code-popover').innerHTML = ['P', 'B'].map(prefix =>
    `<div class="pop-group"><h4>${prefix === 'P' ? '产品侧' : '企业侧'}</h4>` +
    Object.keys(CODES).filter(k => k[0] === prefix).map(k =>
      `<label class="pop-item"><input type="checkbox" data-code="${k}"><span class="pop-code">${k}</span><span>${CODES[k]}</span></label>`).join('') +
    `</div>`).join('');
}
function syncCodeBtn() {
  $('#code-btn').innerHTML = `卡点编码${detail.codes.size ? ` (${detail.codes.size})` : ''}${icon('chevron-down', 14)}`;
}
let popTimer;
function openPopover() {
  const pop = $('#code-popover'), r = $('#code-btn').getBoundingClientRect();
  pop.style.left = Math.max(8, Math.min(r.left, innerWidth - 320)) + 'px';
  pop.style.top = Math.max(8, Math.min(r.bottom + 6, innerHeight - 80)) + 'px';
  pop.hidden = false;
  $('#code-btn').setAttribute('aria-expanded', 'true');
  const show = () => pop.classList.add('show');
  requestAnimationFrame(show);
  setTimeout(show, 50);
}
function closePopover() {
  const pop = $('#code-popover');
  if (pop.hidden) return;
  pop.classList.remove('show');
  $('#code-btn').setAttribute('aria-expanded', 'false');
  clearTimeout(popTimer);
  popTimer = setTimeout(() => pop.hidden = true, 150);
}

/* ---------- 组件菜单（···）与 CSV ---------- */
function csvFunnel() { const { counts, stuck, conv } = funnelStats(); return [['节点','到达企业数','转化率','卡住企业数'], ...NODES.map((v, i) => [v, counts[i], conv(i), stuck[i]])]; }
function csvLeads() {
  const profile = filtered.filter(c => c.level >= 2).length;
  return [['指标','数值'], ['线索有效率', pct(filtered.filter(c => ['高','中'].includes(c['意向等级'])).length, profile)],
    ...['高','中','低'].map(k => [k + '意向', filtered.filter(c => c['意向等级'] === k).length + ' 家'])];
}
function csvDaily() { return [['日期','新进入企业数'], ...DAYS.map((d, i) => [d, dailyNewCounts()[i]])]; }
function csvEvents() {
  if (!data.hasEvents) return [['事件量', '暂无事件数据']];
  const names = new Set(filtered.map(c => c['企业名称']));
  const events = data.events.filter(e => names.has(e['企业名称']));
  return [['日期','事件数'], ...DAYS.map(d => [d, events.filter(e => (e['事件时间'] || '').slice(0, 10) === d).length]), ['合计', events.length], ['日均', (events.length / 15).toFixed(1)]];
}
function csvOppDots() {
  const cats = [...new Set(filtered.flatMap(c => c.cats))].map(k => [k, filtered.filter(c => c.cats.includes(k)).length]).sort((a, b) => b[1] - a[1]);
  return [['推荐类目','企业数'], ...cats];
}
function csvInsight() { return [['洞察','数值','说明'], ...insightSlides.map(s => [s.title, s.num.replace(/<[^>]+>/g, ''), s.desc])]; }
function csvBlocks() {
  const p = blockTab === 'P';
  return [['编码','名称','企业数','占比'], ...Object.keys(CODES).filter(k => k[0] === (p ? 'P' : 'B')).map(k => [k, CODES[k], filtered.filter(c => c.codes.includes(k)).length, pct(filtered.filter(c => c.codes.includes(k)).length, filtered.length)])];
}
function csvNurture() {
  const ns = filtered.filter(c => c['长期培育候选'] === '是');
  return [['长期培育候选', ns.length + ' 家'], ['其中已授权', ns.filter(c => c['触达授权'] === '已授权').length + ' 家'], ['其中未授权', ns.filter(c => c['触达授权'] !== '已授权').length + ' 家']];
}
function csvOpps() {
  const cats = [...new Set(filtered.flatMap(c => c.cats))].map(k => [k, filtered.filter(c => c.cats.includes(k))]).sort((a, b) => b[1].length - a[1].length);
  return [['推荐类目','候选企业数','其中已授权','其中已成交'], ...cats.map(([k, cs]) => [k, cs.length, cs.filter(c => c['触达授权'] === '已授权').length, cs.filter(c => c['成交服务']).length])];
}
function csvDetail() {
  const head = ['企业名称','最远节点','卡点编码','主要缺口','推荐类目','意向等级','触达授权','长期培育候选','最近使用时间'];
  const rows = filtered.map(c => [name(c), cleanNode(c['最远节点']), c['卡点编码'], c['主要缺口'], c['推荐类目'], c['意向等级'], c['触达授权'], c['长期培育候选'], c['最近使用时间']]);
  return [head, ...rows];
}
function csvFeatures() {
  const n = filtered.length;
  const names = new Set(filtered.map(c => c['企业名称']));
  return [['功能','使用企业数','使用率','使用次数'], ...FEATURES.map(([label, key, ev]) => {
    const count = filtered.filter(c => c['用过的功能'].includes(key)).length;
    return [label, count, pct(count, n), data.hasEvents ? data.events.filter(e => names.has(e['企业名称']) && ev.includes(e['事件'])).length : '—'];
  })];
}
function csvProcess() {
  if (!data.hasEvents) return [['过程指标', '暂无事件数据']];
  const fmt = ([label, a, b, desc, type]) => [label, b ? (type ? (a / b).toFixed(1) + (type === 'stars' ? ' 星' : '') : pct(a, b)) : '—', desc];
  return [['指标', '数值', '口径'], ...processMetrics().map(fmt)];
}
const MENUS = {
  funnel: () => [
    { label: '查看卡住企业', act: () => { detail.node = NODES[funnelSel]; jump(); } },
    { label: '导出本图数据为 CSV', act: () => downloadCsv('用户旅程漏斗', csvFunnel()) }],
  leads: () => [{ label: '导出本图数据为 CSV', act: () => downloadCsv('线索有效率', csvLeads()) }],
  daily: () => [{ label: '导出本图数据为 CSV', act: () => downloadCsv('每日进入企业', csvDaily()) }],
  events: () => [{ label: '导出本图数据为 CSV', act: () => downloadCsv('事件量', csvEvents()) }],
  'opp-dots': () => [{ label: '导出本图数据为 CSV', act: () => downloadCsv('商机企业', csvOppDots()) }],
  insight: () => [{ label: '导出本图数据为 CSV', act: () => downloadCsv('洞察', csvInsight()) }],
  blocks: () => [{ label: '导出本图数据为 CSV', act: () => downloadCsv('卡点分布', csvBlocks()) }],
  nurture: () => [{ label: '导出本图数据为 CSV', act: () => downloadCsv('长期培育候选', csvNurture()) }],
  opps: () => [{ label: '导出本图数据为 CSV', act: () => downloadCsv('商机分布', csvOpps()) }],
  detail: () => [{ label: '导出本图数据为 CSV', act: () => downloadCsv('企业明细', csvDetail()) }],
  features: () => [{ label: '导出本图数据为 CSV', act: () => downloadCsv('功能使用', csvFeatures()) }],
  process: () => [{ label: '导出本图数据为 CSV', act: () => downloadCsv('过程指标', csvProcess()) }]
};
let menuTimer, openMenuId = null;
function openMenu(btn) {
  const id = btn.dataset.menu, menu = $('#widget-menu');
  const items = typeof MENUS[id] === 'function' ? MENUS[id]() : [];
  menu.innerHTML = items.map((m, i) => `<button type="button" class="menu-item" data-menu-act="${id}:${i}">${m.label}</button>`).join('');
  menu.hidden = false;
  btn.setAttribute('aria-expanded', 'true');
  const r = btn.getBoundingClientRect();
  menu.style.left = Math.max(8, Math.min(r.right - menu.offsetWidth, innerWidth - menu.offsetWidth - 8)) + 'px';
  menu.style.top = Math.min(r.bottom + 8, innerHeight - menu.offsetHeight - 8) + 'px';
  openMenuId = id;
  const show = () => menu.classList.add('show');
  requestAnimationFrame(show);
  setTimeout(show, 50);
}
function closeMenu() {
  const menu = $('#widget-menu');
  if (menu.hidden) return;
  menu.classList.remove('show');
  $$('[data-menu]').forEach(b => b.setAttribute('aria-expanded', 'false'));
  clearTimeout(menuTimer);
  menuTimer = setTimeout(() => menu.hidden = true, 150);
  openMenuId = null;
}

/* ---------- 通用 tooltip（data-tip） ---------- */
function showTooltip(el) {
  const tip = $('#tooltip');
  tip.textContent = el.dataset.tip;
  const r = el.getBoundingClientRect();
  const w = tip.offsetWidth, h = tip.offsetHeight;
  let left = Math.max(8, Math.min(r.left + r.width / 2 - w / 2, innerWidth - w - 8));
  let top = r.top - h - 8;
  if (top < 8) top = r.bottom + 8;
  tip.style.left = left + 'px'; tip.style.top = top + 'px';
  tip.classList.add('show');
}
const hideTooltip = () => $('#tooltip').classList.remove('show');

/* ---------- 顶部提示 ---------- */
let toastTimer;
function showToast(msg) {
  const t = $('#toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 3000);
}

/* ---------- 抽屉 ---------- */
function openDrawer(c) {
  previousFocus = document.activeElement;
  $('#drawer').innerHTML = drawerContent(c);
  $('#drawer-mask').classList.add('open');
  document.body.style.overflow = 'hidden';
  $('#drawer').scrollTop = 0;
  $('#drawer-close').focus();
}
function closeDrawer() {
  const mask = $('#drawer-mask');
  if (!mask.classList.contains('open')) return;
  mask.classList.remove('open');
  document.body.style.overflow = '';
  previousFocus?.focus();
  previousFocus = null;
}
function drawerContent(c) {
  const kv = keys => `<dl class="kv">${keys.map(k => `<dt>${k}</dt><dd class="${k.includes('时间') ? 'nowrap' : ''}">${esc(k === '最远节点' || k === '卡点节点' ? cleanNode(c[k]) : c[k]) || '—'}</dd>`).join('')}</dl>`;
  const gaps = data.gaps.filter(g => g['企业名称'] === c['企业名称']);
  const events = data.events.filter(e => e['企业名称'] === c['企业名称']).sort((a, b) => a['事件时间'].localeCompare(b['事件时间']));
  const follow = ['跟进人','跟进日期','跟进结果','成交服务','成交金额（元）'];
  const intent = c['意向等级'];
  const status = `<div class="status-tags">` +
    `<span class="s-tag ${intent === '高' ? 'hi' : intent === '中' ? 'md' : 'lo'}">意向等级：${esc(intent)}</span>` +
    `<span class="s-tag ${c['触达授权'] === '已授权' ? 'ok' : ''}">触达授权：${esc(c['触达授权'])}</span>` +
    (c['长期培育候选'] === '是' ? '<span class="s-tag pink">培育候选</span>' : '') + `</div>`;
  const gapCols = ['政策名称','对照维度','政策要求','企业现状','状态','差距'];
  let lastDate = '';
  const html = data.hasEvents && events.length ? events.slice(-30).map(e => {
    const t = e['事件时间'];
    const date = `${t.slice(5, 7)}月${t.slice(8, 10)}日`;
    const head = date !== lastDate ? `<div class="tl-date">${date}</div>` : '';
    lastDate = date;
    const fields = e.fields ? Object.entries(e.fields).map(([k, v]) => `${esc(k)}：${esc(typeof v === 'object' ? JSON.stringify(v) : v)}`).join('；') : '';
    return head + `<div class="tl-item"><span class="tl-time">${t.slice(11, 16)}</span><span class="tl-event">${esc(cleanNode(e['节点']))} · ${esc(e['事件'])}</span><small class="tl-fields">${fields}</small></div>`;
  }).join('') : '';
  const services = c['推荐服务（服务名｜服务 ID）'].split('；').filter(Boolean).map(s => {
    const [label, id] = s.split('｜').map(x => x.trim());
    return `<div class="service"><span class="s-name">${esc(label)}</span><span class="s-id">${esc(id)}</span>` +
      (/^\d+$/.test(id) ? `<a class="pill-mini" href="https://www.wdsheng.com/employer/find-service/detail/${id}" target="_blank" rel="noopener noreferrer">在万大圣查看${icon('external', 12)}</a>` : '') + `</div>`;
  }).join('');
  return `
  <div class="drawer-head">
    <button type="button" id="drawer-close" aria-label="关闭">${icon('close', 16)}</button>
    <h2>${esc(name(c))}</h2>
    <p class="d-meta">${['所在地区','所属行业（国标代码）','企业规模','成立年限','企业类型'].map(k => esc(c[k]) || '—').join(' · ')}</p>
    ${status}
  </div>
  <div class="d-section"><h3>概况</h3>${kv(['意向等级','触达授权','长期培育候选','最远节点','卡点节点','画像完整度','会话数','首次使用时间','最近使用时间', ...['专利数','软著数','研发人员','匹配政策数','最高预估通过率','选定政策','选定政策预估通过率','点过服务入口','选过“暂不需要”'].filter(k => k in c)])}</div>
  <div class="d-section"><h3>卡点</h3>${c.codes.map(k => `<div class="code-line">${codeChip(k)}<span>${codeName(k)}</span><span class="side">${k.startsWith('P') ? '产品侧' : '企业侧'}</span></div>`).join('') || '<p>没有卡点记录</p>'}</div>
  <div class="d-section"><h3>条件缺口</h3>${!data.hasGaps ? missing('条件缺口明细') : gaps.length
    ? `<div class="table-scroll"><table class="d-table"><thead><tr>${gapCols.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>${gaps.map(g => `<tr>${gapCols.map(k => k === '状态' ? `<td><span class="${g[k] === '可能不符' ? 'danger' : g[k] === '符合' ? 'success' : 'zero'}">${esc(g[k])}</span></td>` : `<td>${esc(g[k]) || '—'}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`
    : '<p>没有条件缺口记录</p>'}</div>
  <div class="d-section"><h3>推荐服务</h3>${services || '<p>暂无推荐服务</p>'}</div>
  <div class="d-section"><h3>联系方式</h3>${c['触达授权'] === '已授权' ? `<p>${esc(c['联系手机号']) || '—'}</p>` : '<p class="notice">企业未授权服务触达，不可联系</p>'}</div>
  <div class="d-section"><h3>跟进记录</h3>${follow.some(k => c[k]) ? kv(follow) : '<p>业务尚未填写跟进记录</p>'}</div>
  <div class="d-section"><h3>使用轨迹</h3>${!data.hasEvents ? missing('事件明细') : !events.length ? '<p>暂无使用轨迹</p>' : (events.length > 30 ? `<p>共 ${events.length} 条，显示最近 30 条</p>` : '') + html}</div>`;
}

/* ---------- 事件绑定 ---------- */
buildPopover(); syncCodeBtn();
document.addEventListener('click', e => {
  const el = e.target.closest('button,[data-company],a,label');
  if ($('#widget-menu').hidden === false && !e.target.closest('#widget-menu,[data-menu]')) closeMenu();
  if ($('#code-popover').hidden === false && !e.target.closest('#code-popover,#code-btn')) closePopover();
  if (!el) return;
  if (el.hasAttribute('data-import')) { $('#file').click(); return; }
  if (el.dataset.intent !== undefined) {
    const k = el.dataset.intent;
    filters.intents.has(k) ? filters.intents.delete(k) : filters.intents.add(k);
    el.setAttribute('aria-pressed', filters.intents.has(k));
    page = 1; render(); return;
  }
  if (el.dataset.col !== undefined && el.classList.contains('funnel-col')) {
    const i = Number(el.dataset.col);
    if (funnelSel !== i) { funnelSel = i; renderFunnel(); }
    return;
  }
  if (el.dataset.bandStuck !== undefined) { detail.node = el.dataset.bandStuck; jump(); return; }
  if (el.dataset.insight !== undefined) { showInsight(Number(el.dataset.insight)); return; }
  if (el.dataset.menu !== undefined && el.hasAttribute('data-menu')) {
    openMenuId === el.dataset.menu && !$('#widget-menu').hidden ? closeMenu() : openMenu(el);
    return;
  }
  if (el.dataset.menuAct !== undefined) {
    const [id, i] = el.dataset.menuAct.split(':');
    const items = typeof MENUS[id] === 'function' ? MENUS[id]() : [];
    const item = items[Number(i)];
    closeMenu();
    item?.act();
    return;
  }
  if (el.dataset.node) { detail.node = el.dataset.node; jump(); return; }
  if (el.dataset.tab) {
    if (blockTab !== el.dataset.tab) { blockTab = el.dataset.tab; renderBlocks(); }
    return;
  }
  if (el.dataset.block) { detail.codes.add(el.dataset.block); syncCodeBtn(); jump(); return; }
  if (el.dataset.cat) { detail.cat = el.dataset.cat; jump(); return; }
  if (el.dataset.clear) {
    const k = el.dataset.clear;
    if (k === 'region') { $('#region').value = ''; page = 1; render(); }
    else if (k === 'auth') { $('#auth').value = ''; page = 1; render(); }
    else if (k.startsWith('intent:')) { filters.intents.delete(k.slice(7)); $$('[data-intent]').forEach(b => b.setAttribute('aria-pressed', filters.intents.has(b.dataset.intent))); page = 1; render(); }
    else {
      if (k === 'node') detail.node = '';
      else if (k === 'nurture') detail.nurture = false;
      else if (k === 'cat') detail.cat = '';
      else if (k === 'search') $('#search').value = '';
      else if (k.startsWith('code:')) { detail.codes.delete(k.slice(5)); syncCodeBtn(); }
      page = 1; renderDetail();
    }
    return;
  }
  if (el.id === 'clear-all') { reset(); return; }
  if (el.id === 'view-nurture') { detail.nurture = true; jump(); return; }
  if (el.id === 'drawer-close') { closeDrawer(); return; }
  if (el.id === 'code-btn') { $('#code-popover').hidden ? openPopover() : closePopover(); return; }
  if (el.id === 'copy-link') {
    const done = () => showToast('页面链接已复制');
    if (navigator.clipboard?.writeText) navigator.clipboard.writeText(location.href).then(done, () => showToast('复制失败，请手动复制地址'));
    else showToast('复制失败，请手动复制地址');
    return;
  }
  if (el.dataset.company !== undefined && el.tagName === 'TR') { openDrawer(data.companies[Number(el.dataset.company)]); return; }
  if (el.dataset.sort) { sortDir = sortKey === el.dataset.sort ? -sortDir : 1; sortKey = el.dataset.sort; renderDetail(); return; }
  if (el.dataset.page) { page += Number(el.dataset.page); renderDetail(); return; }
});
$('#file').addEventListener('change', e => { if (e.target.files[0]) loadFile(e.target.files[0]); });
$('#sample').onclick = async () => {
  try {
    const res = await fetch('sample/惠企运营数据-示例.xlsx');
    if (!res.ok) throw Error();
    await loadFile(new File([await res.arrayBuffer()], '惠企运营数据-示例.xlsx'));
  } catch {
    $('#error').hidden = false;
    $('#error').textContent = '文件无法读取，请确认是后台导出的 .xlsx 文件';
  }
};
for (const id of ['region', 'auth']) $('#' + id).onchange = () => { page = 1; render(); };
$('#search').oninput = () => { page = 1; renderDetail(); };
$('#code-popover').onchange = e => {
  const k = e.target.dataset.code;
  if (!k) return;
  e.target.checked ? detail.codes.add(k) : detail.codes.delete(k);
  syncCodeBtn(); page = 1; renderDetail();
};
$('#drawer-mask').onclick = e => { if (e.target.id === 'drawer-mask') closeDrawer(); };
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') {
    if (!$('#widget-menu').hidden) closeMenu();
    else if (!$('#code-popover').hidden) closePopover();
    else closeDrawer();
  }
  if (e.key === 'Tab' && $('#drawer-mask').classList.contains('open')) {
    const items = [...$('#drawer').querySelectorAll('button,a[href],input,select')];
    if (!items.length) return;
    const first = items[0], last = items.at(-1);
    if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
    else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
  }
  if (e.target instanceof Element) {
    const col = e.target.closest('.funnel-col');
    if (col && (e.key === 'ArrowLeft' || e.key === 'ArrowRight')) {
      e.preventDefault();
      const i = Number(col.dataset.col) + (e.key === 'ArrowRight' ? 1 : -1);
      if (i >= 0 && i < 8) { funnelSel = i; renderFunnel(); $(`.funnel-col[data-col="${i}"]`)?.focus(); }
      return;
    }
    const tab = e.target.closest('.block-tabs button');
    if (tab && (e.key === 'ArrowLeft' || e.key === 'ArrowRight')) {
      e.preventDefault();
      blockTab = blockTab === 'B' ? 'P' : 'B';
      renderBlocks();
      $(`.block-tabs [data-tab="${blockTab}"]`)?.focus();
      return;
    }
    if (e.key === 'Enter' && e.target.matches('tr[data-company]')) openDrawer(data.companies[Number(e.target.dataset.company)]);
  }
});
/* 漏斗胶囊提示：选中列常显，悬停/聚焦临时切换到该列 */
document.addEventListener('mouseover', e => {
  const col = e.target.closest('.funnel-col');
  if (col) { showFunnelTip(col); return; }
  const t = e.target.closest('[data-tip]');
  if (t) showTooltip(t);
});
document.addEventListener('mouseout', e => {
  if (e.target.closest('.funnel-col')) hideFunnelTipToSel();
  if (e.target.closest('[data-tip]')) hideTooltip();
});
document.addEventListener('focusin', e => {
  const col = e.target.closest('.funnel-col');
  if (col) { showFunnelTip(col); return; }
  const t = e.target.closest('[data-tip]');
  if (t) showTooltip(t);
});
document.addEventListener('focusout', e => {
  if (e.target.closest('.funnel-col')) hideFunnelTipToSel();
  if (e.target.closest('[data-tip]')) hideTooltip();
});
function jump() { page = 1; renderDetail(); $('#detail').scrollIntoView({ behavior: reducedMotion() ? 'auto' : 'smooth', block: 'start' }); }

/* ---------- 拖拽导入 ---------- */
let dragDepth = 0;
document.addEventListener('dragenter', e => { e.preventDefault(); dragDepth++; $('#drag').hidden = false; document.body.classList.add('dragging'); });
document.addEventListener('dragover', e => e.preventDefault());
document.addEventListener('dragleave', e => { e.preventDefault(); if (--dragDepth <= 0) { $('#drag').hidden = true; document.body.classList.remove('dragging'); } });
document.addEventListener('drop', e => {
  e.preventDefault(); dragDepth = 0;
  $('#drag').hidden = true; document.body.classList.remove('dragging');
  if (e.dataTransfer.files[0]) loadFile(e.dataTransfer.files[0]);
});
