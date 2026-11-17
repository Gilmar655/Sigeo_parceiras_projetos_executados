/* SIGEO Parceiras — painel de programações, mapa e histórico de projetos
   Dados: data/base.js (window.SIGEO_DATA). Atualização local via importação de Excel (IndexedDB). */
'use strict';
const $ = id => document.getElementById(id);
const NF = new Intl.NumberFormat('pt-BR');
const num = v => v == null || isNaN(v) ? '—' : NF.format(Math.round(v * 10) / 10);
const pct = v => v == null || isNaN(v) ? '—' : (v * 100).toFixed(1).replace('.', ',') + '%';
const br = d => d ? d.slice(0, 10).split('-').reverse().join('/') : '—';
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const norm = v => String(v ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
const cls = s => 's-' + String(s || '').replace(/\s+/g, '-');
const WD = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'];
const MONTHS = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
const monthLabel = m => m ? MONTHS[+m.slice(5, 7) - 1] + '/' + m.slice(2, 4) : '';
const STATUS = ['Executado', 'Executado Parcial', 'Cancelado', 'Liberado para Execução', 'Pendente de Aprovação', 'Liberado para Documentação', 'Rascunho'];
const SC = { 'Executado': '#16a34a', 'Executado Parcial': '#f59e0b', 'Cancelado': '#dc2626', 'Liberado para Execução': '#2563eb', 'Pendente de Aprovação': '#8b5cf6', 'Liberado para Documentação': '#0891b2', 'Rascunho': '#94a3b8' };
const scol = s => SC[s] || '#64748b';
const PALETTE = ['#1f5eff', '#d8127d', '#0891b2', '#f59e0b', '#16a34a', '#8b5cf6', '#ef4444', '#0ea5e9', '#84cc16', '#64748b', '#f97316', '#14b8a6'];
const OPEN = new Set(['Liberado para Execução', 'Pendente de Aprovação', 'Liberado para Documentação', 'Rascunho']);
// Regras de exclusão: parceira + descrição do projeto (aplicadas também nas importações)
const EXCLUDE = [['Start', 'Manutenção de Redes e Linhas Aéreas'], ['Conecta', 'Poda de Árvores']];
const excluded = (origin, desc) => EXCLUDE.some(([o, d]) => norm(o) === norm(origin) && norm(d) === norm(desc));
function applyExclusions() { const P = BASE.projects || {}; const before = BASE.prog.length; BASE.prog = BASE.prog.filter(r => !excluded(r[10], (P[r[0]] || [])[2] || '')); return before - BASE.prog.length; }
const TITLES = { overview: 'Painel geral', map: 'Mapa de projetos', projects: 'Projetos', records: 'Programações', partners: 'Parceiras', data: 'Importar / Exportar' };

let BASE, records = [], filtered = [], projLoc = new Map(), view = 'overview', charts = {}, map = null, histMap = null, isLocal = false;
const tableState = { projects: { sort: 'n', dir: -1, page: 0, size: 50 }, records: { sort: 'date', dir: -1, page: 0, size: 50 } };
let mapColor = 'status';

/* ---------- Armazenamento local (IndexedDB) ---------- */
const IDB = {
  open() { return new Promise((res, rej) => { const r = indexedDB.open('sigeo-parceiras-v2', 1); r.onupgradeneeded = () => r.result.createObjectStore('kv'); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }); },
  async get(k) { const db = await this.open(); return new Promise((res, rej) => { const q = db.transaction('kv').objectStore('kv').get(k); q.onsuccess = () => res(q.result); q.onerror = () => rej(q.error); }); },
  async set(k, v) { const db = await this.open(); return new Promise((res, rej) => { const t = db.transaction('kv', 'readwrite'); t.objectStore('kv').put(v, k); t.oncomplete = res; t.onerror = () => rej(t.error); }); },
  async del(k) { const db = await this.open(); return new Promise((res, rej) => { const t = db.transaction('kv', 'readwrite'); t.objectStore('kv').delete(k); t.oncomplete = res; t.onerror = () => rej(t.error); }); }
};

function toast(msg, ms = 5000) { const t = $('toast'); t.innerHTML = msg; t.style.display = 'block'; clearTimeout(toast.t); toast.t = setTimeout(() => t.style.display = 'none', ms); }
const tokens = e => [...new Set(String(e || '').toUpperCase().split(/[,;|\s]+/).filter(Boolean))];
const regionFromCode = p => (String(p).match(/A\.([A-Z]{3})\./) || [])[1] || '';
const normRegion = r => { r = String(r || '').toUpperCase().trim(); return r; };

/* ---------- Normalização ---------- */
function build() {
  applyExclusions();
  const P = BASE.projects || {}, C = BASE.coords || {};
  const counts = new Map();
  records = BASE.prog.map((r, i) => {
    const [proj, date, status, type, power, equip, chi, hi, hf, contr, origin, line] = r;
    const info = P[proj] || [];
    const eqs = tokens(equip);
    let ll = null; for (const e of eqs) if (C[e]) { ll = C[e]; break; }
    let pc = counts.get(proj); if (!pc) counts.set(proj, pc = {});
    for (const e of eqs) if (C[e]) pc[e] = (pc[e] || 0) + 1;
    const rec = {
      i, proj, date: date || '', status: status || 'Não informado', type: type || '', power: power || '', equip: equip || '', eqs,
      chi: +chi || 0, hi: hi || '', hf: hf || '', contr: contr || '', origin: origin || '', line,
      region: normRegion(info[4]) || regionFromCode(proj), desc: info[2] || '', rua: info[3] || '', nota: info[0] || '',
      pstatus: info[7] || (info.length ? '' : 'Sem cadastro'), empresa: info[1] || '',
      lat: ll ? ll[0] : null, lng: ll ? ll[1] : null, circ: ll ? ll[2] : '',
      month: (date || '').slice(0, 7), wd: date ? new Date(date + 'T12:00:00').getDay() : null
    };
    rec.key = norm([proj, power, equip, rec.rua, rec.nota, rec.desc, contr, origin].join(' '));
    return rec;
  });
  projLoc = new Map();
  for (const [p, cnt] of counts) {
    let best = null, n = 0; for (const [e, c] of Object.entries(cnt)) if (c > n) { n = c; best = e; }
    if (best) projLoc.set(p, { lat: C[best][0], lng: C[best][1], eq: best, circ: C[best][2] });
  }
  for (const r of records) r.geo = projLoc.has(r.proj);
  buildFilterOptions();
  const dates = records.map(r => r.date).filter(Boolean).sort();
  const src = Object.entries(BASE.prog.reduce((a, r) => (a[r[10]] = (a[r[10]] || 0) + 1, a), {})).map(([k, v]) => `${k}: <b>${num(v)}</b>`).join('<br>');
  $('baseStamp').innerHTML = `<b>${isLocal ? 'Base local importada' : 'Base publicada'}</b><br>${br(dates[0])} a ${br(dates.at(-1))}<br>${src}<br><span style="opacity:.75">Atualizado: ${esc((BASE.meta && BASE.meta.generated || '').replace('T', ' '))}</span>`;
  $('footInfo').textContent = `${num(records.length)} programações · ${num(counts.size)} projetos · ${num(projLoc.size)} projetos georreferenciados`;
}

const FILTERS = { fOrigin: 'origin', fStatus: 'status', fType: 'type', fRegion: 'region', fContractor: 'contr', fDesc: 'desc', fPStatus: 'pstatus' };
function buildFilterOptions() {
  for (const [id, f] of Object.entries(FILTERS)) {
    const el = $(id), old = el.value, c = {};
    for (const r of records) { const v = r[f] || '(vazio)'; c[v] = (c[v] || 0) + 1; }
    let opts = Object.keys(c).sort((a, b) => f === 'status' ? (STATUS.indexOf(a) + 99 * (STATUS.indexOf(a) < 0)) - (STATUS.indexOf(b) + 99 * (STATUS.indexOf(b) < 0)) : a.localeCompare(b, 'pt-BR'));
    el.innerHTML = '<option value="">Todos</option>' + opts.map(v => `<option value="${esc(v)}">${esc(v)} (${num(c[v])})</option>`).join('');
    if (opts.includes(old)) el.value = old;
  }
  const dates = records.map(r => r.date).filter(Boolean).sort();
  $('fFrom').min = $('fTo').min = dates[0] || ''; $('fFrom').max = $('fTo').max = dates.at(-1) || '';
}

function applyFilters() {
  const q = norm($('fSearch').value), from = $('fFrom').value, to = $('fTo').value, geo = $('fGeo').value;
  const terms = q.split(/\s+/).filter(Boolean);
  const sel = Object.entries(FILTERS).map(([id, f]) => [f, $(id).value]).filter(x => x[1]);
  filtered = records.filter(r =>
    (!from || r.date >= from) && (!to || r.date <= to) &&
    sel.every(([f, v]) => (r[f] || '(vazio)') === v) &&
    (!geo || (geo === '1') === r.geo) &&
    (!terms.length || terms.every(t => r.key.includes(t))));
  $('filterCount').textContent = `${num(filtered.length)} de ${num(records.length)} programações`;
  const labels = { fSearch: 'Pesquisa', fFrom: 'De', fTo: 'Até', fOrigin: 'Parceira', fStatus: 'Status', fType: 'Tipo', fRegion: 'Região', fContractor: 'Contratada', fDesc: 'Descrição', fPStatus: 'Status projeto', fGeo: 'Localização' };
  $('activeChips').innerHTML = Object.keys(labels).filter(id => $(id).value).map(id => {
    const v = $(id).tagName === 'SELECT' ? $(id).selectedOptions[0].text.replace(/ \([\d.]+\)$/, '') : (id === 'fFrom' || id === 'fTo' ? br($(id).value) : $(id).value);
    return `<span class="chip">${labels[id]}: ${esc(v)} <button data-clear="${id}" aria-label="Remover filtro">✕</button></span>`;
  }).join('');
  for (const st of Object.values(tableState)) st.page = 0;
  return filtered;
}

/* ---------- Estatísticas ---------- */
function stats(rows) {
  const s = { total: rows.length, exec: 0, part: 0, canc: 0, open: 0, chi: 0, projects: new Set(), geo: new Set() };
  for (const r of rows) {
    if (r.status === 'Executado') s.exec++; else if (r.status === 'Executado Parcial') s.part++; else if (r.status === 'Cancelado') s.canc++; else s.open++;
    s.chi += r.chi; s.projects.add(r.proj); if (r.geo) s.geo.add(r.proj);
  }
  const done = s.exec + s.part + s.canc;
  s.eff = done ? (s.exec + s.part) / done : null;
  s.execRate = s.total ? s.exec / s.total : null; s.partRate = s.total ? s.part / s.total : null; s.cancRate = s.total ? s.canc / s.total : null;
  s.nProj = s.projects.size; s.geoRate = s.nProj ? s.geo.size / s.nProj : null;
  return s;
}
const groupCount = (rows, fn) => { const m = new Map(); for (const r of rows) { const k = fn(r) || '(vazio)'; m.set(k, (m.get(k) || 0) + 1); } return [...m].sort((a, b) => b[1] - a[1]); };

function projectsAgg(rows) {
  const m = new Map();
  for (const r of rows) {
    let p = m.get(r.proj);
    if (!p) m.set(r.proj, p = { code: r.proj, desc: r.desc, region: r.region, rua: r.rua, pstatus: r.pstatus, nota: r.nota, origins: new Set(), contrs: new Set(), n: 0, exec: 0, part: 0, canc: 0, open: 0, chi: 0, first: r.date, last: r.date, lastStatus: r.status, lastKey: '' });
    p.n++; p.chi += r.chi; p.origins.add(r.origin); p.contrs.add(r.contr);
    if (r.status === 'Executado') p.exec++; else if (r.status === 'Executado Parcial') p.part++; else if (r.status === 'Cancelado') p.canc++; else p.open++;
    if (r.date && r.date < p.first) p.first = r.date;
    const k = r.date + ' ' + r.hi; if (k >= p.lastKey) { p.lastKey = k; p.last = r.date; p.lastStatus = r.status; }
  }
  for (const p of m.values()) { const l = projLoc.get(p.code); p.lat = l?.lat ?? null; p.lng = l?.lng ?? null; p.eq = l?.eq || ''; p.origin = [...p.origins].join(', '); p.contr = [...p.contrs].join(' | '); p.geo = l ? 'Sim' : 'Não'; }
  return [...m.values()];
}

/* ---------- Gráficos ---------- */
function chart(id, cfg) {
  if (charts[id]) charts[id].destroy();
  const el = $(id); if (!el || !window.Chart) return;
  charts[id] = new Chart(el, cfg);
}
function destroyCharts() { for (const c of Object.values(charts)) c.destroy(); charts = {}; }
function setupChartDefaults() {
  if (!window.Chart) return;
  Chart.defaults.font.family = '"Segoe UI",Inter,Roboto,system-ui,sans-serif';
  Chart.defaults.color = '#64748b'; Chart.defaults.borderColor = '#eef2f8';
  Chart.defaults.plugins.legend.labels.usePointStyle = true; Chart.defaults.plugins.legend.labels.boxWidth = 8;
  Chart.defaults.plugins.tooltip.backgroundColor = '#0b1d3a'; Chart.defaults.plugins.tooltip.padding = 10; Chart.defaults.plugins.tooltip.cornerRadius = 8;
  Chart.defaults.maintainAspectRatio = false;
}
function clickFilter(selectId, labels) {
  return (e, els) => { if (!els.length) return; const v = labels[els[0].index]; const sel = $(selectId); if ([...sel.options].some(o => o.value === v)) { sel.value = sel.value === v ? '' : v; refresh(); } };
}
const panel = (title, sub, body, tag = '', span = 6) => `<section class="card panel span-${span}"><div class="panel-h"><div><h2>${title}</h2><p>${sub}</p></div>${tag ? `<span class="tag">${tag}</span>` : ''}</div>${body}</section>`;
const canvas = (id, c = '') => `<div class="chart ${c}"><canvas id="${id}"></canvas></div>`;

function kpis(rows) {
  const s = stats(rows);
  const k = (l, v, d, c) => `<div class="card kpi" style="--c:${c}"><div class="l">${l}</div><div class="v">${v}</div><div class="d">${d}</div></div>`;
  return `<div class="kpis">
    ${k('Programações', num(s.total), `${num(s.nProj)} projetos distintos`, '#1f5eff')}
    ${k('Executadas', pct(s.execRate), `${num(s.exec)} programações`, SC['Executado'])}
    ${k('Execução parcial', pct(s.partRate), `${num(s.part)} programações`, SC['Executado Parcial'])}
    ${k('Canceladas', pct(s.cancRate), `${num(s.canc)} programações`, SC['Cancelado'])}
    ${k('Em aberto', num(s.open), 'Liberadas, pendentes e rascunhos', SC['Liberado para Execução'])}
    ${k('Efetividade', pct(s.eff), '(Exec. + Parcial) ÷ concluídas', '#d8127d')}
    ${k('CHI programado', num(s.chi), 'Soma do CHI das programações', '#0891b2')}
    ${k('Georreferenciados', pct(s.geoRate), `${num(s.geo.size)} projetos no mapa`, '#64748b')}
  </div>`;
}

/* ---------- Visões ---------- */
function vOverview(rows) {
  const months = [...new Set(rows.map(r => r.month).filter(Boolean))].sort();
  const top = projectsAgg(rows).sort((a, b) => b.n - a.n).slice(0, 12);
  $('content').innerHTML = kpis(rows) + `<div class="grid">
    ${panel('Situação das programações', 'Clique em uma fatia para filtrar', canvas('cStatus'), 'STATUS', 4)}
    ${panel('Evolução mensal', 'Programações por mês e status', canvas('cMonth'), 'MENSAL', 8)}
    ${panel('Desempenho por parceira', 'Composição por status (%) · clique para filtrar', canvas('cOrigin', 'sm'), 'PARCEIRAS', 6)}
    ${panel('Tipos de intervenção', 'Volume por tipo e status', canvas('cType', 'sm'), 'TIPO', 6)}
    ${panel('Descrições mais programadas', 'Top 10 atividades (cadastro de projetos)', canvas('cDesc'), 'TOP 10', 6)}
    ${panel('Contratadas', 'Programações por contrato registrado', canvas('cContr'), 'CONTRATOS', 6)}
    ${panel('Região', 'Programações por região', canvas('cRegion', 'sm'), 'REGIÃO', 4)}
    ${panel('Dia da semana', 'Distribuição das datas programadas', canvas('cWeek', 'sm'), 'AGENDA', 4)}
    ${panel('Status do projeto', 'Situação no cadastro de projetos', canvas('cPStatus', 'sm'), 'CADASTRO', 4)}
    ${panel('Projetos com mais programações', 'Clique na linha para abrir o histórico completo', projTable(top, true), 'TOP 12', 12)}
  </div>`;
  const st = groupCount(rows, r => r.status).sort((a, b) => STATUS.indexOf(a[0]) - STATUS.indexOf(b[0]));
  chart('cStatus', { type: 'doughnut', data: { labels: st.map(x => x[0]), datasets: [{ data: st.map(x => x[1]), backgroundColor: st.map(x => scol(x[0])), borderWidth: 2, borderColor: '#fff' }] }, options: { cutout: '64%', plugins: { legend: { position: 'bottom' }, tooltip: { callbacks: { label: c => ` ${c.label}: ${num(c.raw)} (${pct(c.raw / rows.length)})` } } }, onClick: clickFilter('fStatus', st.map(x => x[0])) } });
  const stList = STATUS.filter(s => rows.some(r => r.status === s));
  chart('cMonth', { type: 'bar', data: { labels: months.map(monthLabel), datasets: stList.map(s => ({ label: s, backgroundColor: scol(s), data: months.map(m => rows.filter(r => r.month === m && r.status === s).length), borderRadius: 3 })) }, options: { scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true } }, plugins: { legend: { position: 'bottom' } }, onClick: (e, els) => { if (!els.length) return; const m = months[els[0].index]; const last = new Date(+m.slice(0, 4), +m.slice(5, 7), 0).getDate(); $('fFrom').value = m + '-01'; $('fTo').value = m + '-' + String(last).padStart(2, '0'); refresh(); } } });
  stackedPct('cOrigin', rows, 'origin', 'fOrigin');
  const types = groupCount(rows, r => r.type).map(x => x[0]);
  chart('cType', { type: 'bar', data: { labels: types, datasets: stList.map(s => ({ label: s, backgroundColor: scol(s), data: types.map(t => rows.filter(r => r.type === t && r.status === s).length), borderRadius: 3 })) }, options: { scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true } }, plugins: { legend: { display: false } }, onClick: clickFilter('fType', types) } });
  hbar('cDesc', groupCount(rows, r => r.desc).slice(0, 10), '#1f5eff', 'fDesc');
  hbar('cContr', groupCount(rows, r => r.contr).slice(0, 12), '#d8127d', 'fContractor');
  const rg = groupCount(rows, r => r.region);
  chart('cRegion', { type: 'bar', data: { labels: rg.map(x => x[0]), datasets: [{ data: rg.map(x => x[1]), backgroundColor: PALETTE, borderRadius: 6 }] }, options: { plugins: { legend: { display: false } }, scales: { x: { grid: { display: false } } }, onClick: clickFilter('fRegion', rg.map(x => x[0])) } });
  const wd = [1, 2, 3, 4, 5, 6, 0];
  chart('cWeek', { type: 'bar', data: { labels: wd.map(d => WD[d]), datasets: [{ data: wd.map(d => rows.filter(r => r.wd === d).length), backgroundColor: '#0891b2', borderRadius: 6 }] }, options: { plugins: { legend: { display: false } }, scales: { x: { grid: { display: false } } } } });
  const ps = groupCount(projectsAgg(rows), p => p.pstatus);
  chart('cPStatus', { type: 'doughnut', data: { labels: ps.map(x => x[0]), datasets: [{ data: ps.map(x => x[1]), backgroundColor: PALETTE, borderColor: '#fff', borderWidth: 2 }] }, options: { cutout: '58%', plugins: { legend: { position: 'right' }, tooltip: { callbacks: { label: c => ` ${c.label}: ${num(c.raw)} projetos` } } }, onClick: clickFilter('fPStatus', ps.map(x => x[0])) } });
}
function hbar(id, items, color, filterId) {
  chart(id, { type: 'bar', data: { labels: items.map(x => x[0].length > 42 ? x[0].slice(0, 40) + '…' : x[0]), datasets: [{ data: items.map(x => x[1]), backgroundColor: color, borderRadius: 5, barThickness: 16 }] }, options: { indexAxis: 'y', plugins: { legend: { display: false }, tooltip: { callbacks: { title: c => items[c[0].dataIndex][0] } } }, scales: { y: { grid: { display: false } } }, onClick: filterId ? clickFilter(filterId, items.map(x => x[0])) : undefined } });
}
function stackedPct(id, rows, field, filterId) {
  const labels = groupCount(rows, r => r[field]).map(x => x[0]);
  const groups = [['Executado', ['Executado']], ['Executado Parcial', ['Executado Parcial']], ['Cancelado', ['Cancelado']], ['Em aberto', [...OPEN]]];
  chart(id, { type: 'bar', data: { labels, datasets: groups.map(([g, ss]) => ({ label: g, backgroundColor: g === 'Em aberto' ? '#93b4ff' : scol(g), data: labels.map(l => { const rr = rows.filter(r => r[field] === l); return rr.length ? rr.filter(r => ss.includes(r.status)).length / rr.length * 100 : 0; }) })) }, options: { indexAxis: 'y', scales: { x: { stacked: true, max: 100, ticks: { callback: v => v + '%' } }, y: { stacked: true, grid: { display: false } } }, plugins: { legend: { position: 'bottom' }, tooltip: { callbacks: { label: c => ` ${c.dataset.label}: ${c.raw.toFixed(1).replace('.', ',')}%` } } }, onClick: filterId ? clickFilter(filterId, labels) : undefined } });
}

function statusBar(p) {
  const t = p.n || 1;
  return `<div class="bar-cell" title="Exec ${p.exec} · Parcial ${p.part} · Canc ${p.canc} · Aberto ${p.open}"><span style="width:${p.exec / t * 100}%;background:${SC.Executado}"></span><span style="width:${p.part / t * 100}%;background:${SC['Executado Parcial']}"></span><span style="width:${p.canc / t * 100}%;background:${SC.Cancelado}"></span><span style="width:${p.open / t * 100}%;background:#93b4ff"></span></div>`;
}
function projTable(list, compact) {
  if (!list.length) return '<div class="empty">Nenhum projeto no recorte.</div>';
  const st = tableState.projects;
  const cols = [['code', 'Projeto'], ['desc', 'Descrição'], ['region', 'Região'], ['rua', 'Rua'], ['origin', 'Parceira'], ['n', 'Prog.', 1], ['exec', 'Exec.', 1], ['part', 'Parc.', 1], ['canc', 'Canc.', 1], ['bar', 'Composição'], ['first', 'Primeira'], ['last', 'Última'], ['lastStatus', 'Último status'], ['pstatus', 'Status projeto'], ['chi', 'CHI', 1], ['geo', 'Mapa']];
  const head = cols.map(([k, l, n]) => `<th class="${n ? 'num' : ''} ${!compact && st.sort === k ? 'sorted' : ''}" ${compact || k === 'bar' ? '' : `data-sort="${k}"`}>${l}${!compact && st.sort === k ? (st.dir > 0 ? ' ▲' : ' ▼') : ''}</th>`).join('');
  const body = list.map(p => `<tr data-proj="${esc(p.code)}"><td class="code">${esc(p.code)}</td><td title="${esc(p.desc)}">${esc(p.desc || '—')}</td><td>${esc(p.region)}</td><td title="${esc(p.rua)}">${esc(p.rua || '—')}</td><td>${esc(p.origin)}</td><td class="num"><b>${num(p.n)}</b></td><td class="num">${num(p.exec)}</td><td class="num">${num(p.part)}</td><td class="num">${num(p.canc)}</td><td>${statusBar(p)}</td><td>${br(p.first)}</td><td>${br(p.last)}</td><td><span class="pill ${cls(p.lastStatus)}">${esc(p.lastStatus)}</span></td><td>${esc(p.pstatus || '—')}</td><td class="num">${num(p.chi)}</td><td>${p.geo === 'Sim' ? '📍' : '<span class="muted">—</span>'}</td></tr>`).join('');
  return `<div class="table-wrap"><table data-table="projects"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
}
function sortList(list, st) {
  const k = st.sort, d = st.dir;
  return list.sort((a, b) => { const x = a[k], y = b[k]; return (typeof x === 'number' && typeof y === 'number' ? x - y : String(x ?? '').localeCompare(String(y ?? ''), 'pt-BR', { numeric: true })) * d; });
}
function pager(name, total) {
  const st = tableState[name], pages = Math.max(1, Math.ceil(total / st.size));
  return `<div class="pager"><span>${num(total)} itens · página ${st.page + 1} de ${num(pages)}</span>
    <button class="btn sm" data-page="${name}:first" ${st.page ? '' : 'disabled'}>«</button><button class="btn sm" data-page="${name}:prev" ${st.page ? '' : 'disabled'}>‹</button>
    <button class="btn sm" data-page="${name}:next" ${st.page < pages - 1 ? '' : 'disabled'}>›</button><button class="btn sm" data-page="${name}:last" ${st.page < pages - 1 ? '' : 'disabled'}>»</button>
    <select data-size="${name}">${[25, 50, 100, 250].map(n => `<option ${n === st.size ? 'selected' : ''}>${n}</option>`).join('')}</select></div>`;
}
function vProjects(rows) {
  const st = tableState.projects, list = sortList(projectsAgg(rows), st);
  const pages = Math.max(1, Math.ceil(list.length / st.size)); st.page = Math.min(st.page, pages - 1);
  const s = stats(rows);
  $('content').innerHTML = `<div class="mini-kpis"><div><b>${num(list.length)}</b>projetos</div><div><b>${num(s.geo.size)}</b>com localização</div><div><b>${num(list.filter(p => p.n > 1).length)}</b>com reprogramação</div><div><b>${num(list.length ? s.total / list.length : 0)}</b>programações por projeto</div></div>
  <section class="card panel"><div class="toolbar"><div class="grow"><b>Carteira de projetos</b> <span class="muted">· clique em um projeto para ver o histórico de todas as programações</span></div>
  <button class="btn sm" id="expProjects">⬇ Exportar projetos (Excel)</button></div>
  ${projTable(list.slice(st.page * st.size, (st.page + 1) * st.size))}${pager('projects', list.length)}</section>`;
  $('expProjects').onclick = () => exportExcel('projetos');
}
function vRecords(rows) {
  const st = tableState.records, list = sortList([...rows], st);
  const pages = Math.max(1, Math.ceil(list.length / st.size)); st.page = Math.min(st.page, pages - 1);
  const cols = [['date', 'Data'], ['proj', 'Projeto'], ['status', 'Status'], ['type', 'Tipo'], ['power', 'PowerON'], ['equip', 'Equipamentos'], ['hi', 'Início'], ['hf', 'Fim'], ['chi', 'CHI', 1], ['origin', 'Parceira'], ['contr', 'Contratada'], ['region', 'Região'], ['desc', 'Descrição'], ['rua', 'Rua']];
  const head = cols.map(([k, l, n]) => `<th class="${n ? 'num' : ''} ${st.sort === k ? 'sorted' : ''}" data-sort="${k}">${l}${st.sort === k ? (st.dir > 0 ? ' ▲' : ' ▼') : ''}</th>`).join('');
  const body = list.slice(st.page * st.size, (st.page + 1) * st.size).map(r => `<tr data-proj="${esc(r.proj)}"><td>${br(r.date)}</td><td class="code">${esc(r.proj)}</td><td><span class="pill ${cls(r.status)}">${esc(r.status)}</span></td><td>${esc(r.type)}</td><td>${esc(r.power || '—')}</td><td title="${esc(r.equip)}">${esc(r.equip)}</td><td>${esc(r.hi)}</td><td>${esc(r.hf)}</td><td class="num">${num(r.chi)}</td><td>${esc(r.origin)}</td><td title="${esc(r.contr)}">${esc(r.contr)}</td><td>${esc(r.region)}</td><td title="${esc(r.desc)}">${esc(r.desc || '—')}</td><td title="${esc(r.rua)}">${esc(r.rua || '—')}</td></tr>`).join('');
  $('content').innerHTML = `<section class="card panel"><div class="toolbar"><div class="grow"><b>Programações</b> <span class="muted">· clique em uma linha para abrir o histórico do projeto</span></div>
  <button class="btn sm" id="expCsv">⬇ CSV</button><button class="btn sm primary" id="expXls">⬇ Excel do recorte</button></div>
  ${list.length ? `<div class="table-wrap"><table data-table="records"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>${pager('records', list.length)}` : '<div class="empty">Nenhuma programação no recorte.</div>'}</section>`;
  $('expCsv').onclick = exportCsv; $('expXls').onclick = () => exportExcel('recorte');
}
function vPartners(rows) {
  const origins = groupCount(rows, r => r.origin).map(x => x[0]);
  const months = [...new Set(rows.map(r => r.month).filter(Boolean))].sort();
  const trs = origins.map(o => { const rr = rows.filter(r => r.origin === o), s = stats(rr), d = rr.map(r => r.date).filter(Boolean).sort(); const top = groupCount(rr, r => r.type)[0]; return `<tr data-origin="${esc(o)}"><td><b>${esc(o)}</b></td><td class="num">${num(s.total)}</td><td class="num">${num(s.nProj)}</td><td class="num">${pct(s.execRate)}</td><td class="num">${pct(s.partRate)}</td><td class="num">${pct(s.cancRate)}</td><td class="num">${num(s.open)}</td><td class="num"><b>${pct(s.eff)}</b></td><td class="num">${num(s.chi)}</td><td class="num">${pct(s.geoRate)}</td><td>${esc(top ? top[0] : '—')}</td><td>${br(d[0])} – ${br(d.at(-1))}</td></tr>`; }).join('');
  $('content').innerHTML = `<section class="card panel"><div class="panel-h"><div><h2>Comparativo entre parceiras</h2><p>Clique em uma parceira para filtrar todo o painel. Efetividade = (Executado + Parcial) ÷ (Executado + Parcial + Cancelado).</p></div><span class="tag">${origins.length} PARCEIRAS</span></div>
  <div class="table-wrap"><table data-table="partners"><thead><tr><th>Parceira</th><th class="num">Programações</th><th class="num">Projetos</th><th class="num">Executado</th><th class="num">Parcial</th><th class="num">Cancelado</th><th class="num">Em aberto</th><th class="num">Efetividade</th><th class="num">CHI</th><th class="num">Georref.</th><th>Tipo principal</th><th>Período</th></tr></thead><tbody>${trs}</tbody></table></div></section>
  <div class="grid">
   ${panel('Efetividade mensal por parceira', '(Exec. + Parcial) ÷ concluídas, por mês', canvas('pEff', 'tall'), 'TENDÊNCIA', 7)}
   ${panel('Volume mensal por parceira', 'Quantidade de programações', canvas('pVol', 'tall'), 'VOLUME', 5)}
   ${panel('Tipo de intervenção por parceira', 'Quantidade por tipo', canvas('pType'), 'MIX', 6)}
   ${panel('Cancelamentos por parceira e mês', 'Programações canceladas', canvas('pCanc'), 'CANCELADAS', 6)}
  </div>`;
  const col = i => PALETTE[i % PALETTE.length];
  chart('pEff', { type: 'line', data: { labels: months.map(monthLabel), datasets: origins.map((o, i) => ({ label: o, borderColor: col(i), backgroundColor: col(i), tension: .3, data: months.map(m => stats(rows.filter(r => r.origin === o && r.month === m)).eff * 100 || null), spanGaps: true })) }, options: { scales: { y: { min: 0, max: 100, ticks: { callback: v => v + '%' } }, x: { grid: { display: false } } }, plugins: { legend: { position: 'bottom' }, tooltip: { callbacks: { label: c => ` ${c.dataset.label}: ${c.raw == null ? '—' : c.raw.toFixed(1).replace('.', ',') + '%'}` } } } } });
  chart('pVol', { type: 'bar', data: { labels: months.map(monthLabel), datasets: origins.map((o, i) => ({ label: o, backgroundColor: col(i), borderRadius: 3, data: months.map(m => rows.filter(r => r.origin === o && r.month === m).length) })) }, options: { scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true } }, plugins: { legend: { position: 'bottom' } } } });
  const types = groupCount(rows, r => r.type).map(x => x[0]);
  chart('pType', { type: 'bar', data: { labels: origins, datasets: types.map((t, i) => ({ label: t, backgroundColor: col(i + 2), borderRadius: 3, data: origins.map(o => rows.filter(r => r.origin === o && r.type === t).length) })) }, options: { plugins: { legend: { position: 'bottom' } }, scales: { x: { grid: { display: false } } } } });
  chart('pCanc', { type: 'line', data: { labels: months.map(monthLabel), datasets: origins.map((o, i) => ({ label: o, borderColor: col(i), backgroundColor: col(i), tension: .3, data: months.map(m => rows.filter(r => r.origin === o && r.month === m && r.status === 'Cancelado').length) })) }, options: { plugins: { legend: { position: 'bottom' } }, scales: { x: { grid: { display: false } } } } });
}

/* ---------- Mapa ---------- */
function tileLayers() {
  return {
    'Google Maps (ruas)': L.tileLayer('https://{s}.google.com/vt/lyrs=m&hl=pt-BR&x={x}&y={y}&z={z}', { maxZoom: 21, subdomains: ['mt0', 'mt1', 'mt2', 'mt3'], attribution: 'Mapa &copy; Google' }),
    'Google Híbrido (satélite + ruas)': L.tileLayer('https://{s}.google.com/vt/lyrs=y&hl=pt-BR&x={x}&y={y}&z={z}', { maxZoom: 21, subdomains: ['mt0', 'mt1', 'mt2', 'mt3'], attribution: 'Imagens &copy; Google' }),
    'Google Satélite': L.tileLayer('https://{s}.google.com/vt/lyrs=s&x={x}&y={y}&z={z}', { maxZoom: 21, subdomains: ['mt0', 'mt1', 'mt2', 'mt3'], attribution: 'Imagens &copy; Google' }),
    'Mapa claro': L.tileLayer('https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png', { maxZoom: 20, attribution: '&copy; OpenStreetMap &copy; CARTO' }),
    'OpenStreetMap': L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 19, attribution: '&copy; OpenStreetMap' }),
    'Satélite': L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', { maxZoom: 19, attribution: 'Tiles &copy; Esri' })
  };
}
function vMap(rows, focus) {
  const list = projectsAgg(rows), withLoc = list.filter(p => p.lat != null);
  const origins = [...new Set(rows.map(r => r.origin))].sort();
  const oc = Object.fromEntries(origins.map((o, i) => [o, PALETTE[i % PALETTE.length]]));
  const legend = mapColor === 'status' ? STATUS.filter(s => list.some(p => p.lastStatus === s)).map(s => `<span><i class="dot" style="background:${scol(s)}"></i>${s}</span>`).join('') : origins.map(o => `<span><i class="dot" style="background:${oc[o]}"></i>${esc(o)}</span>`).join('');
  $('content').innerHTML = `<section class="card panel"><div class="toolbar"><div class="grow"><b>${num(withLoc.length)}</b> projetos no mapa <span class="muted">· ${num(list.length - withLoc.length)} sem coordenada do equipamento · posição = equipamento mais programado do projeto</span></div>
   <div class="seg"><button data-color="status" class="${mapColor === 'status' ? 'on' : ''}">Cor: último status</button><button data-color="origin" class="${mapColor === 'origin' ? 'on' : ''}">Cor: parceira</button></div>
   <button class="btn sm" id="expMap">⬇ Exportar localização (Excel)</button></div>
   <div class="map-legend" style="margin-bottom:10px">${legend}<span class="muted">· tamanho do ponto = nº de programações</span></div>
   <div class="map-wrap"><div id="map"></div></div></section>`;
  $('expMap').onclick = () => exportExcel('projetos');
  if (!window.L) { $('map').innerHTML = '<div class="empty">Biblioteca de mapa indisponível.</div>'; return; }
  if (map) { map.remove(); map = null; }
  const tl = tileLayers();
  map = L.map('map', { preferCanvas: true, layers: [tl['Google Maps (ruas)']] }).setView([-23.55, -46.55], 11);
  L.control.layers(tl, null, { position: 'topright' }).addTo(map);
  L.control.scale({ imperial: false }).addTo(map);
  const cluster = L.markerClusterGroup({ chunkedLoading: true, maxClusterRadius: 45, spiderfyOnMaxZoom: true, showCoverageOnHover: false, disableClusteringAtZoom: 16 });
  const byCode = {};
  for (const p of withLoc) {
    const color = mapColor === 'status' ? scol(p.lastStatus) : oc[p.origins.values().next().value];
    const m = L.circleMarker([p.lat, p.lng], { radius: Math.min(5 + Math.sqrt(p.n) * 2, 15), color: '#fff', weight: 1.5, fillColor: color, fillOpacity: .9 });
    m.bindPopup(() => popupHtml(p), { maxWidth: 320 });
    m.bindTooltip(p.code, { direction: 'top', offset: [0, -6] });
    cluster.addLayer(m); byCode[p.code] = m;
  }
  map.addLayer(cluster);
  map.on('popupopen', e => { const b = e.popup.getElement().querySelector('[data-hist]'); if (b) b.onclick = () => openHistory(b.dataset.hist); });
  if (withLoc.length) map.fitBounds(L.latLngBounds(withLoc.map(p => [p.lat, p.lng])), { padding: [30, 30], maxZoom: 15 });
  if (focus && byCode[focus]) cluster.zoomToShowLayer(byCode[focus], () => byCode[focus].openPopup());
  setTimeout(() => map && map.invalidateSize(), 150);
}
function popupHtml(p) {
  const row = (a, b) => `<div class="row"><span>${a}</span><b>${b}</b></div>`;
  return `<div class="pop"><h4 class="code">${esc(p.code)}</h4><div class="muted" style="margin-bottom:6px">${esc(p.desc || 'Sem descrição no cadastro')}</div>
   ${row('Rua', esc(p.rua || '—'))}${row('Região', esc(p.region))}${row('Parceira', esc(p.origin))}${row('Programações', num(p.n))}
   ${row('Exec / Parc / Canc', `${p.exec} / ${p.part} / ${p.canc}`)}${row('Última', br(p.last) + ` · <span class="pill ${cls(p.lastStatus)}">${esc(p.lastStatus)}</span>`)}${row('Equipamento', esc(p.eq))}
   <button class="btn primary sm" data-hist="${esc(p.code)}">Ver histórico de programações</button>
   <div style="display:flex;gap:6px;margin-top:6px"><a class="btn sm" style="flex:1;text-align:center;text-decoration:none" target="_blank" rel="noopener" href="https://www.google.com/maps/search/?api=1&query=${p.lat},${p.lng}">Google Maps</a><a class="btn sm" style="flex:1;text-align:center;text-decoration:none" target="_blank" rel="noopener" href="https://waze.com/ul?ll=${p.lat},${p.lng}&navigate=yes">Waze</a><a class="btn sm" style="flex:1;text-align:center;text-decoration:none" target="_blank" rel="noopener" href="https://www.google.com/maps/@?api=1&map_action=pano&viewpoint=${p.lat},${p.lng}">Street View</a></div></div>`;
}

/* ---------- Histórico do projeto ---------- */
function openHistory(code) {
  const rows = records.filter(r => r.proj === code).sort((a, b) => (b.date + b.hi).localeCompare(a.date + a.hi));
  if (!rows.length) return toast('Projeto não encontrado na base.');
  const info = (BASE.projects || {})[code] || [], cols = BASE.projCols || [];
  const s = stats(rows), loc = projLoc.get(code), d = rows.map(r => r.date).filter(Boolean).sort();
  const days = d.length > 1 ? Math.round((new Date(d.at(-1)) - new Date(d[0])) / 864e5) : 0;
  const infoBoxes = cols.map((c, i) => info[i] ? `<div class="info"><small>${esc(c)}</small><b>${esc(c.startsWith('Data') ? br(info[i]) : info[i] === 'True' ? 'Sim' : info[i] === 'False' ? 'Não' : info[i])}</b></div>` : '').join('');
  const eqs = [...new Set(rows.flatMap(r => r.eqs))];
  const tl = rows.map(r => `<li style="--c:${scol(r.status)}"><div class="t-top"><span class="t-date">${br(r.date)}</span><span class="pill ${cls(r.status)}">${esc(r.status)}</span><span class="muted">${esc(r.hi)}–${esc(r.hf)}</span></div>
     <div class="t-meta">Tipo <b>${esc(r.type)}</b> · PowerON <b>${esc(r.power || '—')}</b> · CHI <b>${num(r.chi)}</b><br>Equip.: ${esc(r.equip)}<br>${esc(r.origin)} · ${esc(r.contr)}</div></li>`).join('');
  const table = `<div class="table-wrap" style="max-height:320px"><table><thead><tr><th>#</th><th>Data</th><th>Status</th><th>Tipo</th><th>PowerON</th><th>Equipamentos</th><th>Início</th><th>Fim</th><th class="num">CHI</th><th>Contratada</th><th>Parceira</th></tr></thead><tbody>${rows.map((r, i) => `<tr><td>${rows.length - i}</td><td>${br(r.date)}</td><td><span class="pill ${cls(r.status)}">${esc(r.status)}</span></td><td>${esc(r.type)}</td><td>${esc(r.power || '—')}</td><td>${esc(r.equip)}</td><td>${esc(r.hi)}</td><td>${esc(r.hf)}</td><td class="num">${num(r.chi)}</td><td>${esc(r.contr)}</td><td>${esc(r.origin)}</td></tr>`).join('')}</tbody></table></div>`;
  $('modalBody').innerHTML = `<div class="h-head"><div><span class="tag">HISTÓRICO DO PROJETO</span><h2 class="code" style="font-size:22px;margin-top:8px">${esc(code)}</h2><div class="sub">${[info[2] || 'Sem descrição no cadastro', info[3], info[4]].filter(Boolean).map(esc).join(' · ')}</div></div>
   <div class="h-actions"><button class="btn sm primary" id="hExport">⬇ Exportar histórico</button>${loc ? `<button class="btn sm" id="hMap">⌖ Ver no mapa geral</button><a class="btn sm" target="_blank" rel="noopener" style="text-decoration:none" href="https://www.google.com/maps/search/?api=1&query=${loc.lat},${loc.lng}">Google Maps</a><a class="btn sm" target="_blank" rel="noopener" style="text-decoration:none" href="https://waze.com/ul?ll=${loc.lat},${loc.lng}&navigate=yes">Waze</a><a class="btn sm" target="_blank" rel="noopener" style="text-decoration:none" href="https://www.google.com/maps/@?api=1&map_action=pano&viewpoint=${loc.lat},${loc.lng}">Street View</a>` : ''}<button class="btn sm" id="hLink">🔗 Copiar link</button></div></div>
   <div class="mini-kpis" style="margin-top:14px"><div><b>${num(s.total)}</b>programações</div><div><b style="color:${SC.Executado}">${num(s.exec)}</b>executadas</div><div><b style="color:${SC['Executado Parcial']}">${num(s.part)}</b>parciais</div><div><b style="color:${SC.Cancelado}">${num(s.canc)}</b>canceladas</div><div><b style="color:${SC['Liberado para Execução']}">${num(s.open)}</b>em aberto</div><div><b>${num(s.chi)}</b>CHI total</div><div><b>${br(d[0])}</b>primeira</div><div><b>${br(d.at(-1))}</b>última</div><div><b>${num(days)}</b>dias entre 1ª e última</div><div><b>${num(eqs.length)}</b>equipamentos</div></div>
   <div class="info-grid">${infoBoxes || '<div class="info"><small>Cadastro</small><b>Projeto sem registro na base de projetos</b></div>'}${loc ? `<div class="info"><small>Coordenada (${esc(loc.eq)})</small><b>${loc.lat.toFixed(6)}, ${loc.lng.toFixed(6)}</b></div><div class="info"><small>Circuito</small><b>${esc(loc.circ || '—')}</b></div>` : ''}</div>
   <div class="h-body"><div><h3 style="margin:0 0 8px;font-size:15px">Linha do tempo das programações</h3><ul class="timeline">${tl}</ul></div>
   <div>${loc ? '<h3 style="margin:0 0 8px;font-size:15px">Localização dos equipamentos</h3><div id="histMap"></div>' : '<div class="empty">Sem coordenadas para os equipamentos deste projeto.</div>'}
   <h3 style="margin:16px 0 8px;font-size:15px">Status ao longo do tempo</h3><div class="chart sm"><canvas id="hChart"></canvas></div></div></div>
   <h3 style="margin:18px 0 8px;font-size:15px">Tabela completa</h3>${table}`;
  $('modal').classList.add('open'); $('modal').setAttribute('aria-hidden', 'false');
  history.replaceState(null, '', '#projeto=' + encodeURIComponent(code));
  $('hExport').onclick = () => exportHistory(code, rows);
  $('hLink').onclick = () => { navigator.clipboard?.writeText(location.href).then(() => toast('Link do histórico copiado.')).catch(() => toast(location.href)); };
  if (loc) $('hMap').onclick = () => { closeModal(); go('map', code); };
  const asc = [...rows].reverse();
  if (charts.hChart) charts.hChart.destroy();
  if (window.Chart) charts.hChart = new Chart($('hChart'), { type: 'scatter', data: { datasets: STATUS.filter(st => rows.some(r => r.status === st)).map(st => ({ label: st, backgroundColor: scol(st), pointRadius: 7, data: asc.map((r, i) => r.status === st ? { x: new Date(r.date + 'T12:00:00').getTime(), y: STATUS.indexOf(st) } : null).filter(Boolean) })) }, options: { plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => ` ${new Date(c.raw.x).toLocaleDateString('pt-BR')} · ${c.dataset.label}` } } }, scales: { x: { type: 'linear', ticks: { callback: v => new Date(v).toLocaleDateString('pt-BR', { month: 'short', year: '2-digit' }) } }, y: { min: -.5, max: STATUS.length - .5, afterBuildTicks: ax => { ax.ticks = STATUS.map((_, i) => ({ value: i })); }, ticks: { autoSkip: false, callback: v => STATUS[v] ? STATUS[v].replace('Liberado para', 'Lib.').replace('Pendente de', 'Pend.') : '' } } } } });
  if (loc && window.L) {
    if (histMap) { histMap.remove(); histMap = null; }
    const tl2 = tileLayers();
    histMap = L.map('histMap', { layers: [tl2['Google Maps (ruas)']] });
    L.control.layers(tl2).addTo(histMap);
    const C = BASE.coords || {}, pts = [];
    for (const e of eqs) {
      if (!C[e]) continue; const last = rows.find(r => r.eqs.includes(e)); const n = rows.filter(r => r.eqs.includes(e)).length;
      const ll = [C[e][0], C[e][1]]; pts.push(ll);
      L.circleMarker(ll, { radius: e === loc.eq ? 10 : 7, color: e === loc.eq ? '#0b1d3a' : '#fff', weight: 2, fillColor: scol(last.status), fillOpacity: .95 }).addTo(histMap)
        .bindPopup(`<b class="code">${esc(e)}</b><br>Circuito ${esc(C[e][2] || '—')}<br>${n} programação(ões) · última ${br(last.date)} <span class="pill ${cls(last.status)}">${esc(last.status)}</span>`);
    }
    if (pts.length > 1) histMap.fitBounds(pts, { padding: [30, 30], maxZoom: 17 }); else histMap.setView(pts[0], 17);
    setTimeout(() => histMap && histMap.invalidateSize(), 200);
  }
}
function closeModal() {
  $('modal').classList.remove('open'); $('modal').setAttribute('aria-hidden', 'true');
  if (histMap) { histMap.remove(); histMap = null; }
  if (charts.hChart) { charts.hChart.destroy(); delete charts.hChart; }
  if (location.hash.startsWith('#projeto=')) history.replaceState(null, '', '#' + view);
}

/* ---------- Exportação ---------- */
const stamp = () => new Date().toISOString().slice(0, 16).replace(/[-:T]/g, '').replace(/(\d{8})(\d{4})/, '$1_$2');
const toDate = d => d ? { serial: Date.UTC(+d.slice(0, 4), +d.slice(5, 7) - 1, +d.slice(8, 10)) / 864e5 + 25569 } : '';
const plain = rows => rows.map(r => { const o = {}; for (const k in r) o[k] = r[k] && r[k].serial ? r[k].serial : r[k]; return o; });
const dateCols = rows => rows.length ? Object.keys(rows[0]).filter(k => rows.some(r => r[k] && r[k].serial)) : [];
function progRow(r) {
  const l = projLoc.get(r.proj);
  return { 'Projeto': r.proj, 'Data Programacao': toDate(r.date), 'Status Programacao': r.status, 'Tipo Intervencao': r.type, 'Numero PowerON': r.power, 'Equipamentos': r.equip, 'CHI': r.chi, 'Horario Inicio': r.hi, 'Horario Fim': r.hf, 'Contratada': r.contr, 'Parceira': r.origin, 'Região': r.region, 'Descrição do Projeto': r.desc, 'Rua': r.rua, 'Nota': r.nota, 'Status do Projeto': r.pstatus, 'Latitude (equip. da linha)': r.lat ?? '', 'Longitude (equip. da linha)': r.lng ?? '', 'Latitude projeto': l?.lat ?? '', 'Longitude projeto': l?.lng ?? '', 'Google Maps': l ? `https://www.google.com/maps/search/?api=1&query=${l.lat},${l.lng}` : '' };
}
function projRow(p) {
  return { 'Projeto': p.code, 'Descrição do Projeto': p.desc, 'Região': p.region, 'Rua': p.rua, 'Nota': p.nota, 'Status do Projeto': p.pstatus, 'Parceira(s)': p.origin, 'Contratada(s)': p.contr, 'Programações': p.n, 'Executado': p.exec, 'Executado Parcial': p.part, 'Cancelado': p.canc, 'Em aberto': p.open, 'CHI total': p.chi, 'Primeira programação': toDate(p.first), 'Última programação': toDate(p.last), 'Último status': p.lastStatus, 'Equipamento de referência': p.eq, 'Latitude': p.lat ?? '', 'Longitude': p.lng ?? '', 'Google Maps': p.lat != null ? `https://www.google.com/maps/search/?api=1&query=${p.lat},${p.lng}` : '', 'Waze': p.lat != null ? `https://waze.com/ul?ll=${p.lat},${p.lng}&navigate=yes` : '' };
}
function sheet(src) {
  const dc = dateCols(src.slice(0, 500)), rows = plain(src);
  const ws = XLSX.utils.json_to_sheet(rows);
  if (rows.length) {
    const keys = Object.keys(rows[0]);
    ws['!cols'] = keys.map(k => ({ wch: Math.min(48, Math.max(k.length + 2, ...rows.slice(0, 300).map(r => String(dc.includes(k) ? '00/00/0000' : r[k] ?? '').length + 1))) }));
    ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length, c: keys.length - 1 } }) };
    dc.forEach(k => { const c = keys.indexOf(k); for (let r = 1; r <= rows.length; r++) { const cell = ws[XLSX.utils.encode_cell({ r, c })]; if (cell && cell.t === 'n') cell.z = 'dd/mm/yyyy'; } });
  }
  return ws;
}
function summaryRows(rows) {
  const s = stats(rows), out = [];
  out.push({ Indicador: 'Gerado em', Valor: new Date().toLocaleString('pt-BR') });
  out.push({ Indicador: 'Filtros', Valor: [...$('activeChips').querySelectorAll('.chip')].map(c => c.textContent.replace('✕', '').trim()).join(' | ') || 'Nenhum' });
  [['Programações', s.total], ['Projetos distintos', s.nProj], ['Executado', s.exec], ['Executado Parcial', s.part], ['Cancelado', s.canc], ['Em aberto', s.open], ['% Executado', pct(s.execRate)], ['% Parcial', pct(s.partRate)], ['% Cancelado', pct(s.cancRate)], ['Efetividade (Exec+Parc)/concluídas', pct(s.eff)], ['CHI total', s.chi], ['Projetos georreferenciados', s.geo.size]].forEach(([a, b]) => out.push({ Indicador: a, Valor: b }));
  return out;
}
function crossRows(rows, field, label) {
  return groupCount(rows, r => r[field]).map(([k]) => { const rr = rows.filter(r => (r[field] || '(vazio)') === k), s = stats(rr); return { [label]: k, 'Programações': s.total, 'Projetos': s.nProj, 'Executado': s.exec, 'Parcial': s.part, 'Cancelado': s.canc, 'Em aberto': s.open, '% Executado': pct(s.execRate), '% Cancelado': pct(s.cancRate), 'Efetividade': pct(s.eff), 'CHI': s.chi }; });
}
function exportExcel(kind) {
  if (!window.XLSX) return toast('Biblioteca Excel indisponível.');
  const rows = kind === 'completo' ? records : filtered, wb = XLSX.utils.book_new();
  const projs = projectsAgg(rows).sort((a, b) => b.n - a.n);
  if (kind !== 'projetos') XLSX.utils.book_append_sheet(wb, sheet(rows.map(progRow)), 'Programações');
  XLSX.utils.book_append_sheet(wb, sheet(projs.map(projRow)), 'Projetos');
  XLSX.utils.book_append_sheet(wb, sheet(summaryRows(rows)), 'Resumo');
  XLSX.utils.book_append_sheet(wb, sheet(crossRows(rows, 'origin', 'Parceira')), 'Por parceira');
  XLSX.utils.book_append_sheet(wb, sheet(crossRows(rows, 'month', 'Mês').sort((a, b) => a['Mês'].localeCompare(b['Mês']))), 'Por mês');
  XLSX.utils.book_append_sheet(wb, sheet(crossRows(rows, 'contr', 'Contratada')), 'Por contratada');
  XLSX.utils.book_append_sheet(wb, sheet(crossRows(rows, 'type', 'Tipo Intervenção')), 'Por tipo');
  XLSX.writeFile(wb, `SIGEO_Parceiras_${kind}_${stamp()}.xlsx`, { compression: true });
  toast(`Excel gerado: ${num(rows.length)} programações e ${num(projs.length)} projetos.`);
}
function exportHistory(code, rows) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet([...rows].reverse().map((r, i) => ({ '#': i + 1, ...progRow(r) }))), 'Histórico');
  XLSX.utils.book_append_sheet(wb, sheet(projectsAgg(rows).map(projRow)), 'Projeto');
  XLSX.writeFile(wb, `Historico_${code.replace(/[^\w.-]+/g, '_')}_${stamp()}.xlsx`, { compression: true });
}
function exportCsv() {
  const rows = filtered.map(progRow); if (!rows.length) return toast('Recorte vazio.');
  const keys = Object.keys(rows[0]), q = v => { v = v && v.serial ? br(new Date((v.serial - 25569) * 864e5).toISOString()) : String(v ?? ''); return /[;"\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
  download('﻿' + [keys.join(';'), ...rows.map(r => keys.map(k => q(r[k])).join(';'))].join('\r\n'), `SIGEO_Parceiras_recorte_${stamp()}.csv`, 'text/csv;charset=utf-8');
}
function download(content, name, type) { const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([content], { type })); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500); }
function exportBaseJs() {
  const used = new Set(BASE.prog.map(r => r[0])), eqs = new Set(BASE.prog.flatMap(r => tokens(r[5])));
  const out = { ...BASE, meta: { ...BASE.meta, generated: new Date().toISOString().slice(0, 16) }, projects: Object.fromEntries(Object.entries(BASE.projects).filter(([k]) => used.has(k))), coords: Object.fromEntries(Object.entries(BASE.coords).filter(([k]) => eqs.has(k))) };
  download('window.SIGEO_DATA=' + JSON.stringify(out) + ';\n', 'base.js', 'application/javascript');
  toast('Arquivo <b>base.js</b> gerado. Substitua <b>data/base.js</b> no repositório para publicar.', 8000);
}
function exportTemplate() {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([BASE.progCols.slice(0, 10).concat(['Parceira (opcional)']), ['DMP/A.LES.26.00001', '17/11/2026', 'Liberado para Execução', 'FI', '16900000', '352ET000000', 0, '10:00', '14:00', 'Start Engenharia (Área 8B, 9  e 11) NEC', 'Start']]), 'Dados');
  XLSX.writeFile(wb, 'Modelo_importacao_programacoes.xlsx');
}

/* ---------- Importação ---------- */
const H = h => norm(h).replace(/\s+/g, ' ');
function excelDate(v) {
  if (v === '' || v == null) return '';
  if (typeof v === 'number') { const d = new Date(Math.round((Math.floor(v) - 25569) * 864e5)); return isNaN(d) ? '' : d.toISOString().slice(0, 10); }
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  const s = String(v).trim(); let m;
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return `${m[1]}-${m[2]}-${m[3]}`;
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/))) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  return '';
}
function excelTime(v) {
  if (v === '' || v == null) return '';
  if (typeof v === 'number') { const f = v - Math.floor(v), t = Math.round(f * 1440); return String(Math.floor(t / 60) % 24).padStart(2, '0') + ':' + String(t % 60).padStart(2, '0'); }
  const m = String(v).match(/(\d{1,2}):(\d{2})/); return m ? m[1].padStart(2, '0') + ':' + m[2] : '';
}
const clean = v => String(v ?? '').replace(/\s+/g, ' ').trim();
function guessOrigin(name, contr) {
  const t = norm(name + ' ' + contr);
  for (const o of ['Start', 'Conecta', 'Engelmig', 'Alpitel', 'Btobace', 'PSE']) if (t.includes(norm(o))) return o;
  return clean(contr).split(/\s+/)[0] || 'Outra';
}
async function handleFiles(files) {
  if (!files.length) return;
  if (!window.XLSX) return toast('Biblioteca Excel indisponível.');
  const log = []; const say = m => { log.push(m); const el = $('importLog'); if (el) { el.style.display = 'block'; el.textContent = log.join('\n'); } };
  $('loading').style.display = 'flex'; $('loading').querySelector('span').textContent = 'Lendo planilhas…';
  await new Promise(r => setTimeout(r, 30));
  const newProg = {}; let nProj = 0, nEq = 0;
  try {
    for (const f of files) {
      say(`📄 ${f.name}`);
      const wb = XLSX.read(await f.arrayBuffer(), { type: 'array', cellDates: false, dense: true });
      for (const sn of wb.SheetNames) {
        const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: '' });
        let hi = aoa.slice(0, 15).findIndex(r => r.map(H).includes('projeto') || r.map(H).includes('numero do equipamento'));
        if (hi < 0) { say(`   · aba "${sn}": ignorada (cabeçalho não reconhecido)`); continue; }
        const hs = aoa[hi].map(H), ix = n => hs.indexOf(n), body = aoa.slice(hi + 1).filter(r => r.some(v => v !== ''));
        if (ix('projeto') >= 0 && ix('data programacao') >= 0 && ix('status programacao') >= 0) {
          const g = k => ix(k);
          const po = ix('parceira') >= 0 ? ix('parceira') : ix('parceira (opcional)');
          let n = 0;
          for (const r of body) {
            const proj = clean(r[g('projeto')]).toUpperCase(); if (!proj) continue;
            const contr = clean(r[g('contratada')]);
            const origin = (po >= 0 && clean(r[po])) || guessOrigin(f.name + ' ' + sn, contr);
            (newProg[origin] = newProg[origin] || []).push([proj, excelDate(r[g('data programacao')]), clean(r[g('status programacao')]), clean(r[g('tipo intervencao')]), clean(r[g('numero poweron')]), clean(r[g('equipamentos')]).toUpperCase(), +r[g('chi')] || 0, excelTime(r[g('horario inicio')]), excelTime(r[g('horario fim')]), contr, origin, hi + 2 + n]);
            n++;
          }
          say(`   · aba "${sn}": ${num(n)} programações`);
        } else if (ix('projeto') >= 0 && ix('status do projeto') >= 0) {
          const map = BASE.projCols.map(c => ix(H(c)));
          for (const r of body) { const code = clean(r[ix('projeto')]).toUpperCase(); if (!code) continue; BASE.projects[code] = map.map((j, k) => j < 0 ? '' : (k === 10 ? excelDate(r[j]) : clean(r[j]))); nProj++; }
          say(`   · aba "${sn}": cadastro de projetos (${num(body.length)} linhas)`);
        } else if (ix('numero do equipamento') >= 0 && ix('latitude') >= 0) {
          const e = ix('numero do equipamento'), la = ix('latitude'), lo = ix('longitude'), ci = ix('circuito');
          for (const r of body) { const lat = parseFloat(String(r[la]).replace(',', '.')), lng = parseFloat(String(r[lo]).replace(',', '.')); if (!r[e] || isNaN(lat) || isNaN(lng)) continue; BASE.coords[clean(r[e]).toUpperCase()] = [+lat.toFixed(6), +lng.toFixed(6), ci >= 0 ? clean(r[ci]) : '']; nEq++; }
          say(`   · aba "${sn}": ${num(nEq)} coordenadas de equipamentos`);
        } else say(`   · aba "${sn}": ignorada (colunas obrigatórias ausentes)`);
      }
    }
    const origins = Object.keys(newProg);
    if (origins.length) {
      const mode = $('importMode') ? $('importMode').value : 'replaceOrigin';
      if (mode === 'replaceAll') BASE.prog = [];
      else if (mode === 'replaceOrigin') BASE.prog = BASE.prog.filter(r => !origins.includes(r[10]));
      for (const o of origins) BASE.prog.push(...newProg[o]);
      say(`✔ Programações atualizadas (${mode === 'append' ? 'acrescentadas' : mode === 'replaceAll' ? 'base substituída' : 'substituídas por parceira'}): ${origins.map(o => `${o} ${num(newProg[o].length)}`).join(', ')}`);
    }
    if (nProj) say(`✔ Cadastro de projetos atualizado: ${num(nProj)} projetos`);
    if (nEq) say(`✔ Coordenadas atualizadas: ${num(nEq)} equipamentos`);
    if (!origins.length && !nProj && !nEq) { say('Nenhuma informação reconhecida. Use o modelo de importação.'); return; }
    BASE.meta = { ...BASE.meta, generated: new Date().toISOString().slice(0, 16), importado: true };
    isLocal = true;
    try { await IDB.set('base', BASE); say('💾 Base salva neste navegador.'); } catch (e) { say('⚠ Não foi possível salvar localmente (' + e.message + '). A atualização vale até fechar a página.'); }
    build(); refresh();
    toast('Importação concluída. Os indicadores foram recalculados.');
  } catch (e) { console.error(e); say('✖ Erro: ' + e.message); toast('Erro ao importar: ' + esc(e.message)); }
  finally { $('loading').style.display = 'none'; $('fileInput').value = ''; }
}
function vData() {
  const s = stats(records);
  $('content').innerHTML = `<div class="grid">
   <section class="card panel span-7"><div class="panel-h"><div><h2>Importar planilhas Excel</h2><p>Atualize a base direto no navegador — os gráficos, mapa e históricos são recalculados.</p></div><span class="tag">IMPORTAÇÃO</span></div>
    <div class="drop" id="drop"><div style="font-size:34px">⬆</div><h3>Arraste os arquivos aqui</h3><p class="muted">.xlsx, .xls ou .csv · vários arquivos de uma vez</p>
    <div style="display:flex;gap:8px;justify-content:center;flex-wrap:wrap;align-items:center"><label class="muted" style="font-size:12px">Modo: <select id="importMode" class="btn sm" style="font-weight:500"><option value="replaceOrigin">Substituir a parceira importada</option><option value="append">Acrescentar às programações atuais</option><option value="replaceAll">Substituir toda a base</option></select></label><button class="btn primary" id="pick">Selecionar arquivos</button></div></div>
    <div id="importLog" class="import-log" style="display:none"></div>
    <h3 style="font-size:14px;margin:18px 0 6px">Arquivos reconhecidos automaticamente</h3>
    <ol class="steps"><li><b>Programações SIGEO</b> (Start, Conecta, Engelmig…): colunas <i>Projeto, Data Programacao, Status Programacao, Tipo Intervencao, Numero PowerON, Equipamentos, CHI, Horario Inicio, Horario Fim, Contratada</i>. A parceira é identificada pelo nome do arquivo/aba ou pela contratada.</li>
    <li><b>Cadastro de projetos</b>: colunas <i>Projeto, Nota, Empresa Executante, Descrição do Projeto, Rua, Regiao, Status do Projeto…</i></li>
    <li><b>Equipamentos (Buscapé)</b>: colunas <i>Número do Equipamento, Circuito, Latitude, Longitude</i> — define a localização no mapa.</li></ol>
    <p class="note">A base importada fica salva somente neste navegador. Para publicar para todos, gere o <b>base.js</b> e substitua <code>data/base.js</code> no repositório.</p></section>
   <section class="card panel span-5"><div class="panel-h"><div><h2>Exportar</h2><p>Arquivos Excel com filtros automáticos e resumos.</p></div><span class="tag">EXPORTAÇÃO</span></div>
    <div style="display:flex;flex-direction:column;gap:8px">
     <button class="btn primary" data-exp="recorte">⬇ Excel do recorte filtrado (${num(filtered.length)} programações)</button>
     <button class="btn" data-exp="completo">⬇ Excel da base completa (${num(records.length)} programações)</button>
     <button class="btn" data-exp="projetos">⬇ Excel de projetos com localização</button>
     <button class="btn" data-exp="csv">⬇ CSV do recorte (separador ;)</button>
     <button class="btn" data-exp="modelo">⬇ Modelo de planilha para importação</button>
     <button class="btn magenta" data-exp="basejs">⬇ base.js para publicar no GitHub</button>
     <button class="btn danger" id="restore" ${isLocal ? '' : 'disabled'}>↺ Restaurar base publicada</button></div>
    <div class="info-grid" style="grid-template-columns:1fr 1fr"><div class="info"><small>Programações</small><b>${num(s.total)}</b></div><div class="info"><small>Projetos</small><b>${num(s.nProj)}</b></div><div class="info"><small>Projetos no mapa</small><b>${num(s.geo.size)} (${pct(s.geoRate)})</b></div><div class="info"><small>Origem dos dados</small><b>${isLocal ? 'Importação local' : 'Base publicada'}</b></div></div>
    <p class="note"><b>Fora do escopo do painel</b> (removidos também nas importações): Start — Manutenção de Redes e Linhas Aéreas; Conecta — Poda de Árvores.</p><p class="note">Fontes da base publicada: ${(BASE.meta.arquivos || []).map(esc).join(' · ')}</p></section></div>`;
  $('pick').onclick = () => $('fileInput').click();
  const drop = $('drop');
  ['dragenter', 'dragover'].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', ev => handleFiles([...ev.dataTransfer.files]));
  $('content').querySelectorAll('[data-exp]').forEach(b => b.onclick = () => { const k = b.dataset.exp; k === 'csv' ? exportCsv() : k === 'modelo' ? exportTemplate() : k === 'basejs' ? exportBaseJs() : exportExcel(k); });
  $('restore').onclick = async () => { if (!confirm('Descartar a base importada neste navegador e voltar à base publicada?')) return; try { await IDB.del('base'); } catch (e) { } location.hash = '#data'; location.reload(); };
}

/* ---------- Navegação ---------- */
function render(focus) {
  destroyCharts();
  if (map && view !== 'map') { map.remove(); map = null; }
  $('pageTitle').textContent = TITLES[view];
  document.querySelectorAll('#nav button').forEach(b => b.classList.toggle('active', b.dataset.view === view));
  const rows = filtered;
  ({ overview: vOverview, map: vMap, projects: vProjects, records: vRecords, partners: vPartners, data: vData })[view](rows, focus);
}
function refresh() { applyFilters(); render(); }
function go(v, focus) { view = v; if (!location.hash.startsWith('#projeto=')) history.replaceState(null, '', '#' + v); $('sidebar').classList.remove('open'); render(focus); window.scrollTo({ top: 0 }); }

function bind() {
  $('nav').addEventListener('click', e => { const b = e.target.closest('button[data-view]'); if (b) go(b.dataset.view); });
  $('menuBtn').onclick = () => $('sidebar').classList.toggle('open');
  let deb; $('fSearch').addEventListener('input', () => { clearTimeout(deb); deb = setTimeout(refresh, 250); });
  ['fFrom', 'fTo', 'fGeo', ...Object.keys(FILTERS)].forEach(id => $(id).addEventListener('change', refresh));
  $('clearFilters').onclick = () => { ['fSearch', 'fFrom', 'fTo', 'fGeo', ...Object.keys(FILTERS)].forEach(id => $(id).value = ''); refresh(); };
  $('activeChips').addEventListener('click', e => { const b = e.target.closest('[data-clear]'); if (b) { $(b.dataset.clear).value = ''; refresh(); } });
  $('btnExportTop').onclick = () => exportExcel('recorte');
  $('btnImportTop').onclick = () => $('fileInput').click();
  $('fileInput').onchange = e => handleFiles([...e.target.files]);
  $('modalClose').onclick = closeModal;
  $('modal').addEventListener('click', e => { if (e.target === $('modal')) closeModal(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && $('modal').classList.contains('open')) closeModal(); });
  document.addEventListener('click', e => {
    const h = e.target.closest('[data-hist]'); if (h) { openHistory(h.dataset.hist); return; }
    const tr = e.target.closest('tr[data-proj]'); if (tr && !e.target.closest('#modal')) { openHistory(tr.dataset.proj); return; }
    const po = e.target.closest('tr[data-origin]'); if (po) { $('fOrigin').value = $('fOrigin').value === po.dataset.origin ? '' : po.dataset.origin; refresh(); return; }
    const th = e.target.closest('th[data-sort]'); if (th) { const name = th.closest('table').dataset.table, st = tableState[name]; if (!st) return; st.dir = st.sort === th.dataset.sort ? -st.dir : (['n', 'exec', 'part', 'canc', 'chi', 'date', 'last', 'first'].includes(th.dataset.sort) ? -1 : 1); st.sort = th.dataset.sort; render(); return; }
    const pg = e.target.closest('[data-page]'); if (pg) { const [name, a] = pg.dataset.page.split(':'), st = tableState[name]; const total = name === 'projects' ? new Set(filtered.map(r => r.proj)).size : filtered.length; const last = Math.max(0, Math.ceil(total / st.size) - 1); st.page = a === 'first' ? 0 : a === 'prev' ? Math.max(0, st.page - 1) : a === 'next' ? Math.min(last, st.page + 1) : last; render(); return; }
    const cb = e.target.closest('[data-color]'); if (cb) { mapColor = cb.dataset.color; render(); }
  });
  document.addEventListener('change', e => { const s = e.target.closest('[data-size]'); if (s) { const st = tableState[s.dataset.size]; st.size = +s.value; st.page = 0; render(); } });
}

async function init() {
  setupChartDefaults();
  BASE = window.SIGEO_DATA;
  try { const local = await IDB.get('base'); if (local && local.prog) { BASE = local; isLocal = true; } } catch (e) { }
  if (!BASE) { $('loading').innerHTML = '<p>Base de dados não encontrada (data/base.js).</p>'; return; }
  BASE.projects = BASE.projects || {}; BASE.coords = BASE.coords || {};
  bind(); build(); applyFilters();
  const h = decodeURIComponent(location.hash.slice(1));
  if (TITLES[h]) view = h;
  render();
  if (h.startsWith('projeto=')) openHistory(h.slice(8));
  $('loading').style.display = 'none';
  if (isLocal) toast('Usando a base importada neste navegador. Para voltar à base publicada, acesse <b>Importar / Exportar</b>.', 7000);
}
document.addEventListener('DOMContentLoaded', init);
