/* Omni Command Center: phone app. Reads Tom's Microsoft 365 (Outlook, To Do,
   Calendar, SharePoint "Claude outputs") with his own sign-in. No server. */
(function () {
'use strict';
const C = window.OMNI_CONFIG;
const DEMO = new URLSearchParams(location.search).has('demo');
const GRAPH = 'https://graph.microsoft.com/v1.0';
const SCOPES = ['User.Read', 'Mail.Read', 'Calendars.Read', 'Tasks.ReadWrite', 'Files.Read.All', 'Sites.Read.All'];
const $ = s => document.querySelector(s);
const view = $('#view');

/* ---------------- helpers ---------------- */
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const money = (n, d = 0) => (n == null || isNaN(n)) ? '–' : (n < 0 ? '-$' : '$') + Math.abs(n).toLocaleString('en-US', {maximumFractionDigits: d, minimumFractionDigits: d});
const kmoney = n => (n == null || isNaN(n)) ? '–' : Math.abs(n) >= 1e6 ? money(n / 1e6, 2) + 'M' : Math.abs(n) >= 1e4 ? money(n / 1e3, 0) + 'K' : money(n);
const pct = n => (n == null || n === '' || isNaN(n)) ? '–' : (+n).toFixed(1) + '%';
const num = v => { if (v == null || v === '' || v === '-') return 0; const n = parseFloat(String(v).replace(/[$,]/g, '')); return isNaN(n) ? 0 : n; };
const TZ = 'America/Los_Angeles';
const dfmt = (d, o) => d ? new Date(d).toLocaleString('en-US', Object.assign({timeZone: TZ}, o)) : '';
const day = d => dfmt(d, {weekday: 'short', month: 'numeric', day: 'numeric'});
const time = d => dfmt(d, {hour: 'numeric', minute: '2-digit'});
const when = d => { if (!d) return ''; const t = new Date(d), n = new Date(); const s = (n - t) / 1000;
  if (s < 3600) return Math.max(1, Math.round(s / 60)) + 'm ago'; if (s < 86400) return Math.round(s / 3600) + 'h ago';
  return day(d); };
const YMD = new Intl.DateTimeFormat('en-CA', {timeZone: TZ, year: 'numeric', month: '2-digit', day: '2-digit'});
const ymd = d => YMD.format(new Date(d));
const utc = s => s ? String(s).slice(0, 19) + 'Z' : s;
const todayYMD = () => ymd(new Date());
const daysUntil = s => { if (!s) return null; const a = new Date(todayYMD() + 'T00:00:00'), b = new Date(String(s).slice(0, 10) + 'T00:00:00'); return Math.round((b - a) / 864e5); };
const kpi = (v, l, cls = '') => `<div class="kpi ${cls}"><div class="v">${v}</div><div class="l">${esc(l)}</div></div>`;
const err = e => `<div class="err">${esc(e && e.message || e)}</div>`;
const store = {
  get(k, d) { try { const v = localStorage.getItem('omni.' + k); return v == null ? d : JSON.parse(v); } catch (_) { return d; } },
  set(k, v) { try { localStorage.setItem('omni.' + k, JSON.stringify(v)); } catch (_) {} }
};
function toast(msg) { const d = document.createElement('div'); d.className = 'card good'; d.style.cssText = 'position:fixed;left:12px;right:12px;bottom:90px;z-index:30'; d.textContent = msg; document.body.appendChild(d); setTimeout(() => d.remove(), 2200); }
function sheet(html) { $('#sheetbody').innerHTML = html; $('#sheet').hidden = false; }
$('#sheetclose').onclick = () => { $('#sheet').hidden = true; };
$('#sheet').onclick = e => { if (e.target.id === 'sheet') $('#sheet').hidden = true; };

/* ---------------- auth ---------------- */
let pca = null, account = null;
function clientId() { return C.clientId || store.get('clientId', ''); }
async function initAuth() {
  if (DEMO) return true;
  if (!clientId()) return false;
  pca = new window.msal.PublicClientApplication({
    auth: {clientId: clientId(), authority: 'https://login.microsoftonline.com/' + C.tenant, redirectUri: location.origin + location.pathname},
    cache: {cacheLocation: 'localStorage'}
  });
  await pca.initialize();
  const r = await pca.handleRedirectPromise().catch(() => null);
  account = (r && r.account) || pca.getAllAccounts()[0] || null;
  if (account) pca.setActiveAccount(account);
  return true;
}
async function token() {
  try { return (await pca.acquireTokenSilent({scopes: SCOPES, account})).accessToken; }
  catch (e) { await pca.acquireTokenRedirect({scopes: SCOPES, account}); throw new Error('Signing in…'); }
}

/* ---------------- graph ---------------- */
async function g(path, opt = {}) {
  if (DEMO) return window.MOCK.graph(path, opt);
  const h = Object.assign({Authorization: 'Bearer ' + await token()}, opt.headers || {});
  if (opt.body && typeof opt.body !== 'string') { h['Content-Type'] = 'application/json'; opt.body = JSON.stringify(opt.body); }
  let r;
  for (let i = 0; i < 3; i++) {
    r = await fetch(path.startsWith('http') ? path : GRAPH + path, {method: opt.method || 'GET', headers: h, body: opt.body});
    if (r.status !== 429 && r.status !== 503) break;
    await new Promise(res => setTimeout(res, 1500 * (i + 1)));
  }
  if (!r.ok) { let m = r.status + ''; try { const j = await r.json(); m += ' ' + (j.error && j.error.message || ''); } catch (_) {} throw new Error('Microsoft 365: ' + m); }
  if (opt.raw) return r;
  if (r.status === 204) return null;
  const ct = r.headers.get('content-type') || '';
  return ct.includes('json') ? r.json() : r.text();
}
const enc = p => p.split('/').map(encodeURIComponent).join('/');
const spItem = p => `/drives/${C.driveId}/root:/${enc(p)}:`;
async function spText(p) { const r = await g(spItem(p) + '/content', {raw: true}); return DEMO ? r : r.text(); }
async function spBuf(p) { const r = await g(spItem(p) + '/content', {raw: true}); return DEMO ? r : r.arrayBuffer(); }
function safeJSON(t) { try { return JSON.parse(t); } catch (_) { return JSON.parse(t.replace(/[\u0000-\u001F]/g, c => c === '\n' ? '\\n' : c === '\t' ? '\\t' : ' ')); } }
async function spJSON(p) { return safeJSON(await spText(p)); }
const board = f => C.paths.boards + '/' + f;

/* cache: memory for this session, last good copy kept for offline */
const mem = {};
async function data(key, fn, maxAgeMin = 5) {
  const m = mem[key];
  if (m && (Date.now() - m.t) / 6e4 < maxAgeMin) return m.v;
  try { const v = await fn(); mem[key] = {v, t: Date.now()}; store.set('c.' + key, {v, t: Date.now()}); return v; }
  catch (e) { const old = store.get('c.' + key); if (old) { old.v.__stale = when(old.t); return old.v; } throw e; }
}

/* ---------------- loaders ---------------- */
const L = {
  me: () => data('me', () => g('/me?$select=displayName,mail')),
  calendar: () => data('cal', () => { const s = new Date(); s.setHours(0, 0, 0, 0); const e = new Date(s); e.setDate(e.getDate() + 2);
    return g(`/me/calendarView?startDateTime=${s.toISOString()}&endDateTime=${e.toISOString()}&$top=40&$orderby=start/dateTime&$select=subject,start,end,location,isAllDay,webLink`,
      {headers: {Prefer: 'outlook.timezone="UTC"'}}); }),
  inbox: () => data('inbox', () => g(`/me/mailFolders/inbox/messages?$top=40&$orderby=receivedDateTime desc&$select=subject,from,receivedDateTime,isRead,bodyPreview,webLink,flag,importance`)),
  lists: () => data('lists', () => g('/me/todo/lists?$top=50'), 30),
  tasks: id => data('tasks.' + id, () => g(`/me/todo/lists/${id}/tasks?$filter=status ne 'completed'&$top=200`), 2),
  queue: () => data('queue', () => spJSON(board('cc/inbox.json'))),
  checkins: () => data('checkins', async () => {
    const since = new Date(Date.now() - 14 * 864e5).toISOString();
    const r = await g(`/me/messages?$filter=receivedDateTime ge ${since} and from/emailAddress/address eq '${C.joyce}'&$orderby=receivedDateTime desc&$top=40&$select=subject,receivedDateTime,body,webLink`);
    return (r.value || []).filter(m => /^Daily Check-In/i.test(m.subject || ''));
  }),
  sigma: () => data('sigma', async () => {
    const since = new Date(Date.now() - 42 * 864e5).toISOString();
    const r = await g(`/me/messages?$filter=receivedDateTime ge ${since} and from/emailAddress/address eq '${C.sigma}'&$orderby=receivedDateTime desc&$top=30&$select=subject,receivedDateTime,webLink,hasAttachments`);
    return (r.value || []).filter(m => C.sigmaReports.some(s => (m.subject || '').toLowerCase().includes(s.toLowerCase())));
  }, 15),
  jobBoard: () => data('jobboard', async () => {
    const r = await g(spItem(C.paths.jobBoardPublish) + '/children?$select=name,folder,lastModifiedDateTime&$top=200');
    const folders = (r.value || []).filter(x => x.folder).sort((a, b) => b.name.localeCompare(a.name));
    for (const f of folders) {
      try { const t = await spText(C.paths.jobBoardPublish + '/' + f.name + '/board-data.txt');
        const j = safeJSON(t.replace(/^\s*window\.__BOARD__\s*=\s*/, '').replace(/;\s*$/, '')); j.__folder = f.name; return j; } catch (_) {}
    }
    throw new Error('No board-data.txt in ' + C.paths.jobBoardPublish);
  }, 30),
  permits: () => data('permits', () => spJSON(C.paths.permits), 30),
  bids: () => data('bids', () => spJSON(board('proposal/bids.json')), 10),
  pl: () => data('pl', () => spJSON(board('finance/pl_class.json')), 60),
  plan: () => data('plan', () => spJSON(board('finance/plan.json')), 240),
  followups: () => data('followups', () => spJSON(board('proposal/followups.json')), 10),
  wrap: () => data('wrap', () => spText('Open Construction Projects/Claude outputs/wrapup/wrapup.html'), 30),
  ref: () => data('ref', () => spJSON(board('marketing/ref.json')), 600),
  qbo: () => data('qbolive', async () => { try { return await qboFromMail(); } catch (e) { const j = await spJSON(board('finance/qbo/live.json')); j.source = (j.source || '') + ' [fallback: ' + e.message + ']'; return j; } }, 30),
  score: () => data('score', async () => parseCSV(await spText(board('finance/scorecard/summary.csv'))), 60),
  scoreStatus: () => data('scoreStatus', () => spJSON(board('finance/scorecard/status.json')), 60),
  paymix: () => data('paymix', () => spJSON(board('finance/paymix-pc.json')).catch(() => spJSON(board('finance/paymix.json'))), 240),
  tools: () => data('tools', async () => readXlsx(await spBuf(C.paths.toolAccounts)), 30),
  posts: () => data('posts', () => spJSON(board('content/posts.json')), 15),
  site: () => data('site', () => spJSON(board('content/siteChecks.json')), 30),
  reviews: () => data('reviews', () => spJSON(board('marketing/reviews.json')), 30),
  uniforms: () => data('uniforms', () => spJSON(C.paths.uniformOrders), 30),
  uniformReports: () => data('uniformReports', async () => {
    const r = await g(`/me/messages?$search="uniform order"&$top=15&$select=subject,receivedDateTime,webLink,hasAttachments,from`);
    return (r.value || []).sort((a, b) => b.receivedDateTime.localeCompare(a.receivedDateTime));
  }, 60)
};

function parseCSV(t) {
  const rows = []; let row = [], f = '', q = false;
  for (let i = 0; i < t.length; i++) { const c = t[i];
    if (q) { if (c === '"') { if (t[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c; }
    else if (c === '"') q = true; else if (c === ',') { row.push(f); f = ''; }
    else if (c === '\n' || c === '\r') { if (c === '\r' && t[i + 1] === '\n') i++; row.push(f); rows.push(row); row = []; f = ''; }
    else f += c; }
  if (f || row.length) { row.push(f); rows.push(row); }
  const h = rows.shift() || []; return rows.filter(r => r.length > 1).map(r => Object.fromEntries(h.map((k, i) => [k, r[i]])));
}
let xlsxLoaded = null;
function loadXlsx() { return xlsxLoaded || (xlsxLoaded = new Promise((res, rej) => { const s = document.createElement('script'); s.src = 'lib/xlsx.mini.min.js'; s.onload = res; s.onerror = rej; document.head.appendChild(s); })); }
async function readXlsx(buf) { await loadXlsx(); const wb = XLSX.read(buf, {type: 'array'}); const out = {};
  wb.SheetNames.forEach(n => { out[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], {header: 1, raw: false, defval: ''}); }); return out; }

/* QuickBooks report emails (CC A/R Aging Summary, CC A/P Aging Summary, CC Balance Sheet, daily 5 AM Excel) -> Cash/A/R/A/P */
function qnum(v) { if (v === '' || v == null) return null; const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/[$,]/g, '').replace(/^\((.*)\)$/, '-$1')); return isNaN(n) ? null : n; }
function agingFromRows(rows) {
  const hi = rows.findIndex(r => r.some(c => /^current$/i.test(String(c).trim())));
  if (hi < 0) throw new Error('No aging header in QuickBooks report');
  const head = rows[hi].map(c => String(c).trim()); const cols = head.slice(1);
  const body = rows.slice(hi + 1).filter(r => String(r[0] || '').trim());
  const parents = new Set(body.map(r => String(r[0]).trim()).filter(l => /^Total for /i.test(l)).map(l => l.replace(/^Total for /i, '')));
  const stack = [], top = new Map(); let grand = null;
  for (const r of body) {
    const label = String(r[0]).trim(); const nums = r.slice(1, head.length).map(qnum);
    if (/^TOTAL$/i.test(label)) { grand = nums; continue; }
    if (/^Total for /i.test(label)) { const nm = label.replace(/^Total for /i, ''); const i = stack.lastIndexOf(nm); if (i >= 0) stack.length = i; if (!stack.length) top.set(nm, nums); continue; }
    if (parents.has(label)) { stack.push(label); continue; }
    if (!nums.some(n => n != null)) { continue; }
    if (!stack.length) top.set(label, nums);
  }
  const ti = cols.findIndex(c => /^total$/i.test(c));
  const bucketCols = cols.map((c, i) => i).filter(i => i !== ti);
  const name = c => c.replace(/\s+/g, '').replace(/^CURRENT$/i, 'Current').replace(/^91ANDOVER$/i, '91+').replace(/^(\d+)-(\d+)$/, '$1-$2');
  const buckets = {}; bucketCols.forEach(i => { buckets[name(cols[i])] = Math.round(((grand ? grand[i] : [...top.values()].reduce((s, v) => s + (v[i] || 0), 0)) || 0) * 100) / 100; });
  const people = [...top.entries()].map(([n, v]) => ({name: n, total: Math.round((v[ti] || 0) * 100) / 100, b: v}));
  const total = grand ? grand[ti] : people.reduce((s, p) => s + p.total, 0);
  const check = Math.abs(people.reduce((s, p) => s + p.total, 0) - total) < 1;
  return {total: Math.round(total * 100) / 100, buckets, people, bucketCols, check};
}
function cashFromBS(rows) {
  const lab = r => String(r[0] || '').trim();
  const a = rows.findIndex(r => /^Bank Accounts$/i.test(lab(r))), b = rows.findIndex(r => /^Total Bank Accounts$/i.test(lab(r)));
  const cash = a >= 0 && b > a ? rows.slice(a + 1, b).filter(r => qnum(r[1]) != null).map(r => ({name: lab(r).replace(/^\d{3,5}\s+/, ''), balance: qnum(r[1])})) : [];
  const get = re => { const r = rows.find(r => re.test(lab(r))); return r ? qnum(r[1]) : null; };
  const dep = get(/^Payments to deposit$/i); if (dep) cash.push({name: 'Payments to deposit (undeposited)', balance: dep});
  const asOf = (rows.find(r => /^As of /i.test(lab(r))) || [''])[0];
  return {cashTotal: b >= 0 ? qnum(rows[b][1]) : null, cash, creditCards: get(/^Total Credit Cards$/i), asOfText: String(asOf), workingCapital: (get(/^Total Current Assets$/i) || 0) - (get(/^Total Current Liabilities$/i) || 0)};
}
function qboBuild(ar, ap, bs) {
  const out = {source: 'QuickBooks 5 AM report emails (CC A/R Aging Summary, CC A/P Aging Summary, CC Balance Sheet)'};
  if (bs) Object.assign(out, cashFromBS(bs.rows), {asOf: bs.at});
  if (ar) { const x = agingFromRows(ar.rows); const over = Object.keys(x.buckets).filter(k => /61|91/.test(k));
    const oi = x.bucketCols.filter((i, j) => /61|91/.test(Object.keys(x.buckets)[j]));
    out.ar = {total: x.total, buckets: x.buckets, check: x.check, over60: over.reduce((s, k) => s + x.buckets[k], 0),
      customers: x.people.filter(p => p.total > 0).map(p => ({name: p.name, total: p.total, over60: oi.reduce((s, i) => s + (p.b[i] || 0), 0)})).sort((a, b) => b.total - a.total)};
    out.asOf = out.asOf || ar.at; }
  if (ap) { const x = agingFromRows(ap.rows); const keys = Object.keys(x.buckets);
    const oi = x.bucketCols.filter((i, j) => /31|61|91/.test(keys[j]));
    out.ap = {total: x.total, buckets: x.buckets, check: x.check, over30: keys.filter(k => /31|61|91/.test(k)).reduce((s, k) => s + x.buckets[k], 0),
      vendors: x.people.filter(p => p.total > 0).map(p => ({name: p.name, total: p.total, over30: oi.reduce((s, i) => s + (p.b[i] || 0), 0)})).sort((a, b) => b.total - a.total)}; }
  return out;
}


async function qboFromMail() {
  const since = new Date(Date.now() - 6 * 864e5).toISOString();
  const r = await g(`/me/messages?$filter=receivedDateTime ge ${since} and from/emailAddress/address eq 'quickbooks@notification.intuit.com'&$orderby=receivedDateTime desc&$top=40&$select=subject,receivedDateTime,hasAttachments`);
  const pick = subj => (r.value || []).find(m => (m.subject || '').trim().toLowerCase() === subj.toLowerCase() && m.hasAttachments);
  const load = async subj => { const m = pick(subj); if (!m) return null;
    const a = await g(`/me/messages/${m.id}/attachments?$select=id,name`); const f = (a.value || []).find(x => /\.xlsx$/i.test(x.name)); if (!f) return null;
    const buf = DEMO ? await g('/att/' + f.id) : await (await g(`/me/messages/${m.id}/attachments/${f.id}/$value`, {raw: true})).arrayBuffer();
    const wb = await readXlsx(buf); return {at: m.receivedDateTime, rows: Object.values(wb)[0]}; };
  const [ar, ap, bs] = await Promise.all([load('CC A/R Aging Summary'), load('CC A/P Aging Summary'), load('CC Balance Sheet')]);
  if (!ar && !ap && !bs) throw new Error('No QuickBooks report emails in the last 6 days');
  return qboBuild(ar, ap, bs);
}



/* ---------------- router ---------------- */
const TABS = {
  command: [['today', 'Today'], ['inbox', 'Inbox'], ['drafts', 'Drafts'], ['todo', 'To Do'], ['admin', 'Mgmt To Do'], ['queue', 'Ask Claude'], ['work', 'Work'], ['systems', 'Systems Review'], ['dash', 'Dashboards'], ['seats', 'Seats']],
  construction: [['checkins', 'Daily Check-Ins'], ['wrap', 'Wrap-Up'], ['jobs', 'Job Board'], ['thursday', 'Thursday Reports'], ['leadtime', 'Lead Times'], ['permits', 'Permits'], ['permitdesk', 'Permit Desk']],
  sales: [['due', 'Bids due'], ['followups', 'Follow-ups'], ['submitted', 'Submitted'], ['hold', 'On hold'], ['kit', 'Tools'], ['bidboard', 'Bid Board']],
  finance: [['cash', 'Cash · A/R · A/P'], ['plan', '2027 Plan'], ['pl', 'P&L by dept'], ['score', 'Scorecard'], ['paymix', 'Payment mix'], ['tools', 'Tool Accounts']],
  marketing: [['content', 'Content'], ['site', 'Site health'], ['standout', 'How We Stand Out'], ['uniforms', 'Uniforms'], ['reviews', 'Reviews']]
};
let cur = store.get('route', {tab: 'command', sub: {}});
function sub() { return cur.sub[cur.tab] || TABS[cur.tab][0][0]; }
function go(tab, s) { cur.tab = tab; if (s) cur.sub[tab] = s; store.set('route', cur); render(); }
document.querySelectorAll('#tabs button').forEach(b => b.onclick = () => go(b.dataset.tab));
$('#refresh').onclick = () => { Object.keys(mem).forEach(k => delete mem[k]); render(); };

let seq = 0;
async function render() {
  const my = ++seq; refreshSeat();
  document.querySelectorAll('#tabs button').forEach(b => b.classList.toggle('on', b.dataset.tab === cur.tab));
  $('#sub').innerHTML = TABS[cur.tab].map(([k, l]) => `<button data-s="${k}" class="${k === sub() ? 'on' : ''}">${esc(l)}</button>`).join('');
  $('#sub').querySelectorAll('button').forEach(b => b.onclick = () => go(cur.tab, b.dataset.s));
  view.innerHTML = '<div class="loading">Loading…</div>';
  $('#refresh').classList.add('spin');
  try {
    const html = await V[cur.tab + '_' + sub()]();
    if (my !== seq) return;
    view.innerHTML = html;
    (W[cur.tab + '_' + sub()] || (() => {}))();
  } catch (e) { if (my === seq) view.innerHTML = err(e); }
  if (my === seq) { $('#refresh').classList.remove('spin'); $('#stamp').textContent = time(new Date()); }
}
const staleNote = d => d && d.__stale ? `<div class="muted center">Offline copy from ${esc(d.__stale)}</div>` : '';

/* ---------------- views ---------------- */
const V = {}, W = {};

/* COMMAND */
V.command_today = async () => {
  const [cal, inbox, bids, checks, permits, me] = await Promise.all([
    L.calendar().catch(() => null), L.inbox().catch(() => null), L.bids().catch(() => null),
    L.checkins().catch(() => null), L.permits().catch(() => null), L.me().catch(() => null)]);
  const t = todayYMD();
  const unread = inbox ? (inbox.value || []).filter(m => !m.isRead).length : '–';
  const active = bids ? Object.values(bids.docs || {}).filter(b => b.section === 'active') : [];
  const dueWeek = active.filter(b => { const d = daysUntil(b.due); return d != null && d >= 0 && d <= 7; });
  const late = active.filter(b => { const d = daysUntil(b.due); return d != null && d < 0; });
  const latest = checks && checks[0] ? parseCheckins(checks[0]) : [];
  const flags = latest.filter(c => c.flags.length);
  const needsTom = permits ? Object.values(permits.portals || {}).filter(p => p.needsTom).length : 0;
  const evs = cal ? (cal.value || []).filter(e => ymd(utc(e.start.dateTime)) === t) : [];
  const engSeat = /engineering/.test(SEAT.primary);
  let h = (engSeat ? `<div class="card flag" onclick="OMNI.go('command','seats')"><div class="t">Takeover on: engineering@ is the active Claude seat</div><div class="s">Tap for seat details or to hand back.</div></div>` : '') + `<h2>Good ${new Date().getHours() < 12 ? 'morning' : 'afternoon'}${me ? ', ' + esc((me.displayName || '').split(' ')[0]) : ''}<small>${esc(day(new Date()))}</small></h2>
  <div class="kpis">${kpi(unread, 'Unread in Inbox', unread > 0 ? 'hot' : '')}${kpi(dueWeek.length, 'Bids due in 7 days', dueWeek.length ? 'hot' : '')}
  ${kpi(late.length, 'Active bids past due', late.length ? 'warn' : '')}${kpi(flags.length + '/' + latest.length, 'Check-ins flagged', flags.length ? 'warn' : '')}</div>`;
  h += `<h2>Today<small>${evs.length} on the calendar</small></h2>`;
  h += evs.length ? evs.map(e => `<div class="card"><div class="row"><div><div class="t">${esc(e.subject)}</div><div class="s">${e.isAllDay ? 'All day' : time(utc(e.start.dateTime)) + ' – ' + time(utc(e.end.dateTime))}${e.location && e.location.displayName ? ' · ' + esc(e.location.displayName) : ''}</div></div></div></div>`).join('') : '<div class="empty">Nothing on the calendar today.</div>';
  if (flags.length) h += `<h2>Field needs you<small>${esc(latest.date || '')}</small></h2>` + flags.map(checkinCard).join('');
  if (dueWeek.length) h += `<h2>Bids due this week</h2>` + dueWeek.sort((a, b) => a.due.localeCompare(b.due)).map(bidCard).join('');
  if (late.length) h += `<div class="card flag"><div class="t">${late.length} active bid${late.length > 1 ? 's' : ''} past the due date</div><div class="s">Still marked active on the Bid Board</div><div class="acts"><button class="btn" onclick="OMNI.go('sales','due')">Review in Sales</button></div></div>`;
  if (needsTom) h += `<div class="card flag"><div class="t">${needsTom} permit portal${needsTom > 1 ? 's' : ''} need your login</div><div class="acts"><button class="btn" onclick="OMNI.go('construction','permits')">Open Permits</button></div></div>`;
  return h;
};

V.command_todo = async () => {
  const lists = (await L.lists()).value || [];
  const main = lists.find(l => l.displayName.toLowerCase() === C.todoMain.toLowerCase()) || lists.find(l => l.wellknownListName === 'defaultList') || lists[0];
  const pick = store.get('todoList', main && main.id);
  const list = lists.find(l => l.id === pick) || main;
  return todoView(list, lists, 'todoList');
};
async function todoView(list, lists, pickKey, isAdmin) {
  if (!list) return '<div class="empty">No To Do lists found.</div>';
  const tasks = ((await L.tasks(list.id)).value || []).sort((a, b) => (a.dueDateTime ? a.dueDateTime.dateTime : '9').localeCompare(b.dueDateTime ? b.dueDateTime.dateTime : '9'));
  const t = todayYMD();
  const od = tasks.filter(x => x.dueDateTime && x.dueDateTime.dateTime.slice(0, 10) < t);
  const td = tasks.filter(x => x.dueDateTime && x.dueDateTime.dateTime.slice(0, 10) === t);
  let h = `<div class="kpis">${kpi(od.length, 'Overdue', od.length ? 'bad' : '')}${kpi(td.length, 'Due today', td.length ? 'hot' : '')}${kpi(tasks.length, 'Open')}${kpi(esc(list.displayName), 'List')}</div>`;
  if (pickKey) h += `<div class="seg">${lists.map(l => `<button class="${l.id === list.id ? 'on' : ''}" onclick="OMNI.pickList('${pickKey}','${l.id}')">${esc(l.displayName)}</button>`).join('')}</div>`;
  if (isAdmin) h += `<div class="acts"><button class="btn" onclick="OMNI.adminReset()">Change admin list</button></div>`;
  h += `<div class="addrow"><input id="newTask" placeholder="Add a task (due today)"><button class="btn pri" onclick="OMNI.addTask('${list.id}')">Add</button></div>`;
  h += tasks.map(x => { const due = x.dueDateTime ? x.dueDateTime.dateTime.slice(0, 10) : ''; const cls = due && due < t ? 'bad' : due === t ? 'flag' : '';
    return `<div class="card ${cls} todo"><button class="check" aria-label="Done" onclick="OMNI.done(this,'${list.id}','${x.id}')"></button><div class="grow"><div class="t">${esc(x.title)}</div><div class="s">${due ? 'Due ' + esc(day(due + 'T12:00:00Z')) : 'No date'}${x.importance === 'high' ? ' · ★' : ''}</div>${x.body && x.body.content && x.body.content.trim() ? `<div class="p muted">${esc(x.body.content.trim().slice(0, 240))}</div>` : ''}</div></div>`; }).join('') || '<div class="empty">All clear.</div>';
  return h;
}
V.command_inbox = async () => {
  const r = await L.inbox(); const ms = r.value || [];
  return staleNote(r) + `<h2>Inbox<small>${ms.filter(m => !m.isRead).length} unread</small></h2>` + ms.map(m => `<div class="card ${m.isRead ? '' : 'good'}" onclick="OMNI.mail('${m.id}')">
    <div class="row"><div class="t">${esc(m.from && m.from.emailAddress ? m.from.emailAddress.name : '')}</div><div class="s">${esc(when(m.receivedDateTime))}</div></div>
    <div class="${m.isRead ? '' : 't'}" style="font-size:14px">${esc(m.subject)}</div><div class="s">${esc((m.bodyPreview || '').slice(0, 140))}</div></div>`).join('');
};
/* CONSTRUCTION */
const CHECK_FLAGS = [['Question?', 'Question Detail', 'Question'], ['Needs Materials?', 'Materials Detail', 'Materials'], ['GC & Client Issues?', 'GC & Client Detail', 'GC/Client'], ['Any Holdups?', 'Holdup Detail', 'Holdup'], ['Safety Issues?', 'Safety Details', 'Safety']];
function parseCheckins(m) {
  const doc = new DOMParser().parseFromString((m.body && m.body.content) || '', 'text/html');
  const cards = [...doc.querySelectorAll('div')].filter(d => d.parentElement === doc.body || (d.parentElement && d.parentElement.tagName === 'BODY'));
  const out = [];
  (cards.length ? cards : [doc.body]).forEach(c => {
    const f = {};
    c.querySelectorAll('strong').forEach(s => { const k = s.textContent.replace(/:\s*$/, '').trim(); const v = (s.parentElement.textContent || '').replace(s.textContent, '').trim(); f[k] = v; });
    if (!f.Title) return;
    const flags = CHECK_FLAGS.filter(([q]) => /^y/i.test(f[q] || '')).map(([q, d, l]) => ({label: l, detail: f[d] || ''}));
    out.push({title: f.Title, by: f['Created By'], date: f['Created (Date)'], flags, f});
  });
  out.date = (m.subject || '').replace(/^Daily Check-In\s*-\s*/i, '');
  return out;
}
function checkinCard(c) {
  return `<div class="card ${c.flags.length ? (c.flags.some(f => f.label === 'Safety') ? 'bad' : 'flag') : ''}"><div class="row"><div class="t">${esc(c.title)}</div><div class="s">${esc(c.by || '')}</div></div>
  ${c.flags.map(f => `<div class="p"><span class="pill ${f.label === 'Safety' ? 'r' : 'y'}">${esc(f.label)}</span> ${esc(f.detail)}</div>`).join('') || '<div class="s">No issues</div>'}</div>`;
}
V.construction_checkins = async () => {
  const msgs = await L.checkins();
  const seen = new Set(); const days = [];
  msgs.forEach(m => { const k = m.subject; if (seen.has(k)) return; seen.add(k); days.push(m); });
  if (!days.length) return '<div class="empty">No Daily Check-In emails in the last 14 days.</div>';
  return staleNote(msgs) + days.slice(0, 7).map((m, i) => { const cs = parseCheckins(m); const fl = cs.filter(c => c.flags.length);
    return `<h2>${esc(cs.date)}<small>${cs.length} check-ins · ${fl.length} flagged</small></h2>` + [...fl, ...cs.filter(c => !c.flags.length)].map(checkinCard).join(''); }).join('');
};
V.construction_jobs = async () => {
  const b = await L.jobBoard(); const jobs = (b.jobs || []).filter(j => !/closed|complete/i.test(j.status || ''));
  const filt = store.get('jobDept', 'All'); const depts = ['All', ...new Set(jobs.map(j => j.department).filter(Boolean))];
  const list = jobs.filter(j => filt === 'All' || j.department === filt).map(j => { const r = (j.job && j.job.runs || []).slice(-1)[0] || {}; return {j, r, t: r.targets || {}, a: r.actuals || {}}; });
  const behind = list.filter(x => x.t.hours && x.a.hours > x.t.hours);
  let h = `<div class="kpis">${kpi(list.length, 'Open jobs')}${kpi(behind.length, 'Over budget hours', behind.length ? 'warn' : '')}${kpi(list.reduce((s, x) => s + (x.a.hours || 0), 0).toFixed(0), 'Hours to date')}${kpi(kmoney(list.reduce((s, x) => s + (x.a.totalCost || 0), 0)), 'Cost to date')}</div><div class="muted">Job Board publish ${esc(b.__folder || '')}</div>
  <div class="seg">${depts.map(d => `<button class="${d === filt ? 'on' : ''}" onclick="OMNI.set('jobDept','${esc(d)}')">${esc(d)}</button>`).join('')}</div>`;
  h += list.sort((x, y) => (y.a.hours / (y.t.hours || 1)) - (x.a.hours / (x.t.hours || 1))).map(({j, r, t, a}) => { const m = j.job && j.job.meta || {}; const used = t.hours ? a.hours / t.hours : 0;
    const fr = a.fieldReports || {};
    return `<div class="card ${used > 1 ? 'bad' : used > .85 ? 'flag' : ''}"><div class="row"><div><div class="t">${esc(m.jobNumber || j.key)} ${esc(m.jobName || '')}</div><div class="s">${esc(m.customer || '')} · ${esc(m.manager || '')}</div></div><span class="pill">${esc(j.status || '')}</span></div>
    <div class="s" style="margin-top:6px">Hours ${(a.hours || 0).toFixed(0)} / ${t.hours ? Math.round(t.hours) : '–'} · Cost ${kmoney(a.totalCost)} / ${kmoney(t.totalCost)} · Margin ${pct(a.marginPct)} (bid ${pct(t.marginPct)})</div>
    <div class="bar ${used > 1 ? 'over' : ''}"><i style="width:${Math.min(100, used * 100).toFixed(0)}%"></i></div>
    <div class="s">${fr.expected ? `Field reports ${fr.submitted}/${fr.expected} · ` : ''}${t.targetDate ? 'Target ' + esc(t.targetDate) : ''} · as of ${esc(r.label || r.capturedAt || '')}</div>
    ${j.reportUrl ? `<div class="acts"><a class="btn" href="${esc(j.reportUrl)}" target="_blank" rel="noopener">KPI report</a></div>` : ''}</div>`; }).join('');
  return staleNote(b) + h;
};
V.construction_thursday = async () => {
  const ms = await L.sigma();
  if (!ms.length) return '<div class="empty">No Sigma Thursday reports in the last 6 weeks.</div>';
  const groups = {}; ms.forEach(m => { const k = C.sigmaReports.find(s => (m.subject || '').toLowerCase().includes(s.toLowerCase())); (groups[k] = groups[k] || []).push(m); });
  return staleNote(ms) + Object.entries(groups).map(([k, arr]) => `<h2>${esc(k)}<small>every Thursday 7 PM</small></h2>` + arr.slice(0, 4).map((m, i) =>
    `<div class="card ${i === 0 ? 'good' : ''}"><div class="row"><div class="t">${esc(day(m.receivedDateTime))}</div><div class="s">${esc(time(m.receivedDateTime))}${i === 0 ? ' · latest' : ''}</div></div>
    <div class="acts"><button class="btn pri" onclick="OMNI.report('${m.id}')">Open report</button><a class="btn" href="${esc(m.webLink)}" target="_blank" rel="noopener">Email</a></div></div>`).join('')).join('');
};
V.construction_permits = async () => {
  const d = await L.permits(); const ps = Object.entries(d.permits || {}).map(([id, p]) => Object.assign({id}, p));
  const open = ps.filter(p => !/closed|final|cancel|void|assumed/i.test((p.portalStatus || '') + ' ' + (p.stage || '')));
  const insp = Object.values(d.inspections || {}).filter(i => i.status === 'required' || i.status === 'scheduled' || i.status === 'partial');
  const portals = Object.values(d.portals || {}).filter(p => p.needsTom);
  const fees = open.filter(p => p.feesDue > 0);
  let h = `<div class="kpis">${kpi(open.length, 'Open permits')}${kpi(insp.length, 'Inspections open')}${kpi(portals.length, 'Logins needed', portals.length ? 'warn' : '')}${kpi(money(fees.reduce((s, p) => s + p.feesDue, 0)), 'Fees due', fees.length ? 'warn' : '')}</div><div class="muted">Updated ${esc(d.updatedAt || '')}</div>`;
  if (portals.length) h += `<h2>Need your login</h2>` + portals.map(p => `<div class="card flag"><div class="t">${esc(p.jurisdiction)}</div><div class="s">${esc((p.notes || '').split('|')[0])}</div><div class="acts"><a class="btn" href="${esc(p.url)}" target="_blank" rel="noopener">Open portal</a></div></div>`).join('');
  h += `<h2>Open permits</h2>` + open.sort((a, b) => (b.lastChangeDate || '').localeCompare(a.lastChangeDate || '')).map(p => `<div class="card ${p.feesDue > 0 ? 'flag' : ''}"><div class="row"><div><div class="t">${esc(p.project || p.permitNumber)}</div><div class="s">${esc(p.jurisdiction)} · ${esc(p.permitNumber)}</div></div><span class="pill">${esc(p.portalStatus || p.stage || '')}</span></div>
    ${p.lastChangeNote ? `<div class="p">${esc(p.lastChangeNote)}</div>` : ''}${p.feesDue > 0 ? `<div class="s">Fees due ${money(p.feesDue, 2)}</div>` : ''}</div>`).join('');
  if (insp.length) h += `<h2>Inspections still to pass</h2>` + insp.map(i => `<div class="card"><div class="t">${esc(i.type)}</div><div class="s">${esc(i.jurisdiction)} · ${esc(i.permitNumber)} · ${esc(i.status)}${i.date ? ' · ' + esc(i.date) : ''}</div></div>`).join('');
  return h;
};

/* SALES */
function bidCard(b) {
  const d = daysUntil(b.due); const tag = d == null ? '' : d < 0 ? `${-d}d late` : d === 0 ? 'Due today' : `in ${d}d`;
  return `<div class="card ${d != null && d < 0 ? 'bad' : d != null && d <= 2 ? 'flag' : ''}"><div class="row"><div><div class="t">${esc(b.title || b.sp_name || b.id)}</div>
  <div class="s">${esc(b.estimator || '')}${b.gc ? ' · ' + esc(b.gc) : ''}${b.city ? ' · ' + esc(b.city) : ''}</div></div><div class="center"><div class="amt">${b.due ? esc(day(b.due + 'T12:00:00Z')) : 'No date'}</div><div class="s">${tag}</div></div></div>
  ${b.walk_date ? `<div class="s">Walk ${esc(b.walk_date)}</div>` : ''}${b.note ? `<div class="p muted">${esc(String(b.note).slice(0, 180))}</div>` : ''}
  <div class="acts">${b.sp_url ? `<a class="btn" href="${esc(b.sp_url)}" target="_blank" rel="noopener">Bid folder</a>` : ''}${b.pandadoc && b.pandadoc.draft_id ? `<a class="btn" href="https://app.pandadoc.com/a/#/documents/${esc(b.pandadoc.draft_id)}" target="_blank" rel="noopener">PandaDoc</a>` : ''}</div></div>`;
}
async function bidList(section, sort) {
  const b = await L.bids(); const who = store.get('bidWho', 'All');
  let arr = Object.values(b.docs || {}).filter(x => x.section === section);
  const people = ['All', ...new Set(arr.map(x => x.estimator).filter(Boolean))];
  arr = arr.filter(x => who === 'All' || x.estimator === who).sort(sort);
  return staleNote(b) + `<div class="seg">${people.map(p => `<button class="${p === who ? 'on' : ''}" onclick="OMNI.set('bidWho','${esc(p)}')">${esc(p)}</button>`).join('')}</div><div class="muted">${arr.length} bids · board updated ${esc(when(b.updatedAt))}</div>` + arr.map(bidCard).join('');
}
V.sales_due = () => bidList('active', (a, b) => (a.due || '9999').localeCompare(b.due || '9999'));
V.sales_submitted = () => bidList('submitted', (a, b) => (b.due || '').localeCompare(a.due || ''));
V.sales_hold = () => bidList('on_hold', (a, b) => (b.updated_at || '').localeCompare(a.updated_at || ''));

/* FINANCE */
V.finance_cash = async () => {
  let q; try { q = await L.qbo(); } catch (e) { q = null; }
  if (!q) return `<div class="card flag"><div class="t">Waiting on the QuickBooks export</div><div class="s">Cash, A/R and A/P come from Claude outputs/boards/finance/qbo/live.json, which your PC writes from the QuickBooks report emails. It isn't there yet.</div></div>`;
  const ar = q.ar || {}, ap = q.ap || {};
  let h = `<div class="muted">QuickBooks ${esc(q.asOfText || '')} · emailed ${esc(dfmt(q.asOf, {month: 'numeric', day: 'numeric', hour: 'numeric', minute: '2-digit'}))}${(q.ar && q.ar.check === false) || (q.ap && q.ap.check === false) ? ' · <span style="color:var(--warn)">totals check off, verify in QuickBooks</span>' : ''}</div>
  <div class="kpis">${kpi(kmoney(q.cashTotal), 'Cash in bank', 'hot')}${kpi(kmoney(ar.total), 'A/R open')}${kpi(kmoney(ar.over60), 'A/R over 60 days', ar.over60 > 0 ? 'warn' : '')}${kpi(kmoney(ap.total), 'A/P open')}</div>`;
  if (q.cash && q.cash.length) h += `<h2>Bank accounts</h2><div class="tbl"><table>${q.cash.map(c => `<tr><td>${esc(c.name)}</td><td class="n">${money(c.balance, 2)}</td></tr>`).join('')}</table></div>`;
  const aging = (o) => o.buckets ? `<div class="tbl"><table><tr>${Object.keys(o.buckets).map(k => `<th class="n">${esc(k)}</th>`).join('')}</tr><tr>${Object.values(o.buckets).map(v => `<td class="n">${kmoney(v)}</td>`).join('')}</tr></table></div>` : '';
  h += `<h2>A/R aging<small>who to chase</small></h2>${aging(ar)}` + (ar.customers || []).slice(0, 25).map(c => `<div class="card ${c.over60 > 0 ? 'flag' : ''}"><div class="row"><div class="t">${esc(c.name)}</div><div class="amt">${money(c.total)}</div></div>${c.over60 > 0 ? `<div class="s">${money(c.over60)} over 60 days</div>` : ''}</div>`).join('');
  h += `<h2>A/P aging</h2>${aging(ap)}` + (ap.vendors || []).slice(0, 25).map(c => `<div class="card"><div class="row"><div class="t">${esc(c.name)}</div><div class="amt">${money(c.total)}</div></div>${c.over30 > 0 ? `<div class="s">${money(c.over30)} over 30 days</div>` : ''}</div>`).join('');
  return h;
};
V.finance_pl = async () => {
  const pl = await L.pl(); const per = store.get('plPer', 'mtd'); const p = pl.periods[per];
  const cls = p.classes.filter(c => c !== 'Not specified');
  const sec = (s) => p.rows.filter(r => r.section === s);
  const sum = (rows, c) => rows.reduce((t, r) => t + (r.amounts[c] || 0), 0);
  const inc = sec('Income'), cogs = sec('Cost of Goods Sold');
  const cols = [...cls, 'Total'];
  const ni = p.check && p.check.netIncome || {};
  let h = `<div class="seg">${[['mtd', 'This month'], ['lm', 'Last month'], ['ytd', 'Year to date']].map(([k, l]) => `<button class="${k === per ? 'on' : ''}" onclick="OMNI.set('plPer','${k}')">${l}</button>`).join('')}</div>
  <div class="muted">${esc(p.dateRange)} · QuickBooks P&L by Class</div>
  <div class="kpis">${kpi(kmoney(sum(inc, 'Total')), 'Income', 'hot')}${kpi(kmoney(sum(inc, 'Total') - sum(cogs, 'Total')), 'Gross profit')}${kpi(pct((sum(inc, 'Total') - sum(cogs, 'Total')) / (sum(inc, 'Total') || 1) * 100), 'Gross margin')}${kpi(kmoney(ni.Total), 'Net income', ni.Total < 0 ? 'bad' : '')}</div>`;
  h += `<h2>By department</h2>` + cols.map(c => { const i = sum(inc, c), cg = sum(cogs, c), gp = i - cg;
    return `<div class="card ${c === 'Total' ? 'good' : ''}"><div class="row"><div class="t">${esc(c)}</div><div class="amt">${kmoney(i)}</div></div><div class="s">COGS ${kmoney(cg)} · Gross profit ${kmoney(gp)} (${pct(i ? gp / i * 100 : null)})${ni[c] != null ? ' · Net ' + kmoney(ni[c]) : ''}</div>
    <div class="bar"><i style="width:${Math.max(0, Math.min(100, i ? gp / i * 100 : 0)).toFixed(0)}%"></i></div></div>`; }).join('');
  return h;
};
V.finance_score = async () => {
  const rows = await L.score(); const st = await L.scoreStatus().catch(() => null);
  const per = store.get('scPer', 'this_month'); const pers = [...new Map(rows.map(r => [r.period, r.period_label])).entries()];
  const sel = rows.filter(r => r.period === per && r.part === 'Total');
  const tot = sel.reduce((s, r) => s + num(r.revenue), 0), fh = sel.reduce((s, r) => s + num(r.field_hours), 0), bh = sel.reduce((s, r) => s + num(r.billed_hours), 0);
  let h = `<div class="seg">${pers.map(([k, l]) => `<button class="${k === per ? 'on' : ''}" onclick="OMNI.set('scPer','${k}')">${esc(l)}</button>`).join('')}</div>
  <div class="muted">${st && st.ranAt ? 'Refreshed ' + esc(when(st.ranAt)) : ''}</div>
  <div class="kpis">${kpi(kmoney(tot), 'Revenue', 'hot')}${kpi(fh.toFixed(0), 'Field hours')}${kpi(bh.toFixed(0), 'Billed hours')}${kpi(pct(fh ? bh / fh * 100 : null), 'Efficiency', fh && bh / fh < .6 ? 'warn' : '')}</div>`;
  h += sel.map(r => { const parts = rows.filter(x => x.period === per && x.department === r.department && x.part !== 'Total'); const eff = num(r.efficiency_pct);
    return `<div class="card"><div class="row"><div class="t">${esc(r.department)}</div><div class="amt">${kmoney(num(r.revenue))}</div></div>
    <div class="s">${r.jobs_completed && r.jobs_completed !== '0' ? r.jobs_completed + ' jobs (avg ' + kmoney(num(r.avg_job_price)) + ')' : ''}${r.projects_completed && r.projects_completed !== '0' ? ' · ' + r.projects_completed + ' projects (avg ' + kmoney(num(r.avg_project_price)) + ')' : ''}</div>
    <div class="s">Hours ${num(r.field_hours).toFixed(0)} field · ${num(r.billed_hours).toFixed(0)} billed · efficiency ${pct(r.efficiency_pct)}</div><div class="bar ${eff && eff < 50 ? 'over' : ''}"><i style="width:${Math.min(100, eff)}%"></i></div>
    ${parts.map(p => `<div class="s">↳ ${esc(p.part)}: ${kmoney(num(p.revenue))} · eff ${pct(p.efficiency_pct)}</div>`).join('')}</div>`; }).join('');
  return h;
};
V.finance_paymix = async () => {
  const d = await L.paymix(); const rows = d.rows || [];
  const T = k => rows.reduce((s, r) => s + (r[k] || 0), 0);
  const fees = T('card_fees') + T('ach_fees');
  const upd = (d.meta || []).find(m => m.label === 'Updated');
  let h = `<div class="muted">${upd ? esc(upd.value) : ''}</div><div class="kpis">${kpi(kmoney(T('collected')), 'Collected 2026', 'hot')}${kpi(money(fees), 'Processing fees')}${kpi(pct(fees / (T('collected') || 1) * 100), 'Blended cost')}${kpi(pct(T('card') / (T('collected') || 1) * 100), 'Paid by card')}</div>`;
  h += rows.map(r => { const f = r.card_fees + r.ach_fees; return `<div class="card"><div class="row"><div class="t">${esc(r.department)}</div><div class="amt">${kmoney(r.collected)}</div></div>
    <div class="s">Card ${kmoney(r.card)} · Payabli ACH ${kmoney(r.payabli_ach)} · Bank ACH ${kmoney(r.bank_ach)} · Check ${kmoney(r.check_other)}</div>
    <div class="s">Fees ${money(f, 2)} = ${pct(f / (r.collected || 1) * 100)} of collected</div></div>`; }).join('');
  return h;
};
V.finance_tools = async () => {
  const wb = await L.tools(); const acc = wb.Accounts || [];
  const hi = acc.findIndex(r => r[0] === 'Employee'); const H = acc[hi] || []; const col = n => H.findIndex(x => String(x).startsWith(n));
  const rows = acc.slice(hi + 1).filter(r => r[0] && /yes/i.test(r[col('Active')]));
  const c = {exp: col('EXPECTED'), rec: col('Recorded'), chk: col('Check'), wait: col('Awaiting'), av: col('Available'), wk: col('Weeks'), rate: col('Weekly rate'), cap: col('Cap')};
  const owed = rows.reduce((s, r) => s + num(r[c.exp]), 0); const review = rows.filter(r => !/ok/i.test(r[c.chk] || 'OK'));
  const carrying = rows.filter(r => num(r[c.exp]) > 0).sort((a, b) => num(b[c.exp]) - num(a[c.exp]));
  const rq = wb.Requests || []; const rh = rq.findIndex(r => r[0] === 'Req #'); const RH = rq[rh] || [];
  const reqs = rq.slice(rh + 1).filter(r => r[0]); const stI = RH.indexOf('Status');
  const openReq = reqs.filter(r => !/purchased|denied|closed|done/i.test(r[stI] || ''));
  let h = `<div class="kpis">${kpi(money(owed), 'Owed to Omni (expected)', 'hot')}${kpi(carrying.length, 'People carrying a balance')}${kpi(openReq.length, 'Open requests', openReq.length ? 'warn' : '')}${kpi(review.length, 'Flagged REVIEW', review.length ? 'bad' : '')}</div>
  <div class="muted">From Tool_Account_Master.xlsx (Accounting/AP). $750 cap, $50/week payroll deduction.</div>`;
  if (openReq.length) h += `<h2>Open requests</h2>` + openReq.map(r => `<div class="card flag"><div class="row"><div class="t">${esc(r[RH.indexOf('Employee')])}: ${esc(r[RH.indexOf('Item')])}</div><div class="amt">${esc(r[RH.indexOf('Price ($)')])}</div></div><div class="s">#${esc(r[0])} · ${esc(r[stI])} · ${esc(r[RH.indexOf('Date submitted')])}</div></div>`).join('');
  h += `<h2>Accounts<small>${rows.length} active</small></h2>` + [...carrying, ...rows.filter(r => num(r[c.exp]) <= 0)].map(r => { const bal = num(r[c.exp]), cap = num(r[c.cap]) || 750;
    return `<div class="card ${!/ok/i.test(r[c.chk] || 'OK') ? 'bad' : bal > cap ? 'flag' : ''}"><div class="row"><div class="t">${esc(r[0])}</div><div class="amt">${money(bal, 2)}</div></div>
    <div class="s">Available ${esc(r[c.av])} · ${esc(r[c.rate])}/wk${num(r[c.wk]) ? ' · ' + esc(r[c.wk]) + ' weeks left' : ''}${num(r[c.wait]) ? ' · awaiting ' + esc(r[c.wait]) : ''}${!/ok/i.test(r[c.chk] || 'OK') ? ' · ' + esc(r[c.chk]) : ''}</div>
    <div class="bar ${bal > cap ? 'over' : ''}"><i style="width:${Math.min(100, bal / cap * 100).toFixed(0)}%"></i></div></div>`; }).join('');
  h += `<div class="acts"><a class="btn" href="${C.links.toolPolicy}" target="_blank" rel="noopener">Program policy</a></div>`;
  return h;
};

/* MARKETING */
V.marketing_content = async () => {
  const d = await L.posts(); const ps = Object.entries(d.docs || {}).map(([id, p]) => Object.assign({id}, p));
  const pend = ps.filter(p => p.status === 'pending'), appr = ps.filter(p => p.status === 'approved').sort((a, b) => (a.date || '').localeCompare(b.date || ''));
  const posted = ps.filter(p => p.status === 'posted').sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  const card = p => `<div class="card ${p.status === 'pending' ? 'flag' : p.status === 'approved' ? 'good' : ''}"><div class="row"><div><div class="t">${esc(p.title || p.id)}</div><div class="s">${esc(p.channel || '')} · ${esc(p.date || '')} ${esc(p.time || '')}${p.pillar ? ' · ' + esc(p.pillar) : ''}</div></div><span class="pill">${esc(p.status)}</span></div>
    ${p.tomNote ? `<div class="p muted">Your note: ${esc(p.tomNote)}</div>` : ''}<div class="acts">${p.url || p.link ? `<a class="btn" href="${esc(p.url || p.link)}" target="_blank" rel="noopener">Open</a>` : ''}<button class="btn" onclick="OMNI.post('${esc(p.id)}')">Read</button></div></div>`;
  OMNI._posts = Object.fromEntries(ps.map(p => [p.id, p]));
  return `<div class="kpis">${kpi(pend.length, 'Waiting on you', pend.length ? 'hot' : '')}${kpi(appr.length, 'Approved, scheduled')}${kpi(posted.length, 'Posted')}${kpi(posted.filter(p => daysUntil(p.date) > -8).length, 'Posted last 7 days')}</div>`
    + (pend.length ? '<h2>Needs your OK</h2>' + pend.map(card).join('') : '') + (appr.length ? '<h2>Scheduled</h2>' + appr.map(card).join('') : '') + '<h2>Recently posted</h2>' + posted.slice(0, 12).map(card).join('')
    + `<div class="acts"><a class="btn" href="${C.links.claudeBoard}" target="_blank" rel="noopener">Approve / edit on the full board</a></div>`;
};
V.marketing_site = async () => {
  const d = await L.site(); const days = Object.values(d.docs || {}).sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  const s = days[0]; if (!s) return '<div class="empty">No site checks yet.</div>'; const k = s.kpis || {};
  let h = `<div class="muted">omnimechusa.com · check of ${esc(s.date)}</div><div class="kpis">${kpi(k.speedHomeMob ?? '–', 'Phone speed (home)', k.speedHomeMob < 90 ? 'warn' : '')}${kpi(k.gscClicks28 ?? '–', 'Google clicks, 28d')}${kpi(k.ga4Sessions28 ?? '–', 'Site visits, 28d')}${kpi((k.leadsReal ?? '–') + '', 'Real leads (window)', k.leadsReal ? 'hot' : '')}</div>`;
  h += `<div class="kpis" style="margin-top:8px">${kpi(k.gscImpr28 ?? '–', 'Google impressions')}${kpi(k.aiReferrals28 ?? '–', 'AI referrals')}${kpi(k.socialViews28 ?? '–', 'Social views')}${kpi(k.indexed ?? '–', 'Pages indexed')}</div>`;
  if ((s.leads || []).length) h += '<h2>Leads</h2>' + s.leads.map(l => `<div class="card good"><div class="t">${esc(l.name || l.channel || 'Lead')}</div><div class="s">${esc(l.channel || '')} · ${esc(l.source || '')}</div>${l.summary || l.text ? `<div class="p">${esc(l.summary || l.text)}</div>` : ''}</div>`).join('');
  if ((s.issues || []).length) h += '<h2>Issues</h2>' + s.issues.map(i => `<div class="card flag"><div class="t">${esc(i.text)}</div><div class="s">Fix: ${esc(i.fix || '')} (${esc(i.who || '')})</div></div>`).join('');
  if ((s.fixed || []).length) h += '<h2>Fixed</h2>' + s.fixed.map(f => `<div class="card good"><div class="p">${esc(f)}</div></div>`).join('');
  if ((s.blocked || []).length) h += '<h2>Blocked</h2>' + s.blocked.map(f => `<div class="card"><div class="p">${esc(f)}</div></div>`).join('');
  return h;
};
V.marketing_reviews = async () => {
  let d; try { d = await L.reviews(); } catch (e) { return `<div class="card"><div class="t">No review file yet</div><div class="s">Reviews show here from Claude outputs/boards/marketing/reviews.json.</div></div>`; } const rs = Array.isArray(d) ? d : d.reviews || Object.values(d.docs || d.items || {});
  const toDo = rs.filter(r => !r.replied && !r.reply && !/answered|replied|done/i.test(r.status || ''));
  return `<div class="kpis">${kpi(rs.length, 'Reviews')}${kpi(toDo.length, 'To answer', toDo.length ? 'hot' : '')}${kpi(rs.length ? (rs.reduce((s, r) => s + num(r.rating || r.stars), 0) / rs.length).toFixed(1) + '★' : '–', 'Average')}${kpi(esc(d.updatedAt ? when(d.updatedAt) : ''), 'Updated')}</div>`
    + rs.slice(0, 40).map(r => `<div class="card ${toDo.includes(r) ? 'flag' : ''}"><div class="row"><div class="t">${esc(r.author || r.name || r.reviewer || 'Review')}</div><div class="amt">${'★'.repeat(num(r.rating || r.stars))}</div></div><div class="s">${esc(r.source || r.platform || '')} · ${esc(r.date || '')}</div>${r.text || r.comment ? `<div class="p">${esc(r.text || r.comment)}</div>` : ''}${r.reply ? `<div class="p muted">Reply: ${esc(r.reply)}</div>` : ''}</div>`).join('');
};

/* ---------------- more views (Wrap-Up, Follow-ups, Tools, Plan, Positioning) ---------------- */
V.construction_wrap = async () => {
  const html = await L.wrap();
  W.construction_wrap = () => {
    const f = document.getElementById('wrapframe'); if (!f) return;
    f.src = URL.createObjectURL(new Blob([String(html)], {type: 'text/html'}));
    f.onload = () => { try { f.style.height = (f.contentDocument.documentElement.scrollHeight + 20) + 'px'; } catch (_) { f.style.height = '4000px'; } };
  };
  return staleNote(html) + `<div class="muted">Latest Construction Wrap-Up (SharePoint wrapup.html)</div><iframe id="wrapframe" sandbox="allow-same-origin allow-popups" style="width:100%;border:0;border-radius:12px;background:#fff;min-height:70vh;margin-top:8px"></iframe>`;
};

V.sales_followups = async () => {
  const d = await L.followups(); const who = store.get('fuWho', 'All');
  let arr = Object.values(d.docs || {}).filter(x => x && !x._deleted);
  const people = ['All', ...new Set(arr.map(x => x.estimator).filter(Boolean))];
  arr = arr.filter(x => who === 'All' || x.estimator === who);
  const open = arr.filter(x => /due|ask_next|waiting/.test(x.status || '')).sort((a, b) => (a.next_date || '9').localeCompare(b.next_date || '9'));
  const done = arr.filter(x => !open.includes(x));
  const total = open.reduce((s, x) => s + (x.amount || 0), 0);
  const card = x => { const late = x.next_date && daysUntil(x.next_date) < 0; const dr = x.draft && x.draft.state === 'open' ? x.draft : null;
    return `<div class="card ${x.status === 'due' ? (late ? 'bad' : 'flag') : x.status === 'replied' ? 'good' : ''}"><div class="row"><div><div class="t">${esc(x.job || x.bid_id)}</div><div class="s">${esc(x.company || '')}${x.contact && x.contact.name ? ' · ' + esc(x.contact.name) : ''} · ${esc(x.estimator || '')}</div></div><div class="center"><div class="amt">${x.amount ? kmoney(x.amount) : ''}</div><div class="s">${esc(x.status || '')}</div></div></div>
    <div class="s">Sent ${esc(x.sent_date || '–')} · last touch ${esc(x.last_touch && x.last_touch.date || '–')} · step ${esc(x.step || '–')}${x.next_date ? ' · next ' + esc(x.next_date) + (late ? ' (late)' : '') : ''}</div>
    ${dr ? `<div class="p muted">${esc(dr.body)}</div>` : ''}<div class="acts">${dr && dr.web_link ? `<a class="btn pri" href="${esc(dr.web_link)}" target="_blank" rel="noopener">Open draft in Outlook</a>` : ''}${x.contact && x.contact.email ? `<a class="btn" href="mailto:${esc(x.contact.email)}">Email</a>` : ''}</div></div>`; };
  return staleNote(d) + `<div class="seg">${people.map(p => `<button class="${p === who ? 'on' : ''}" onclick="OMNI.set('fuWho','${esc(p)}')">${esc(p)}</button>`).join('')}</div>
  <div class="kpis">${kpi(open.filter(x => x.status === 'due').length, 'Follow-ups due', 'hot')}${kpi(kmoney(total), 'Open proposal $')}${kpi(arr.filter(x => x.status === 'replied').length, 'Replied')}${kpi(esc(when(d.updatedAt)), 'Updated')}</div>`
    + '<h2>To follow up</h2>' + (open.map(card).join('') || '<div class="empty">Nothing due.</div>') + (done.length ? '<h2>Other</h2>' + done.map(card).join('') : '');
};

/* Tools: private copies in SharePoint Claude outputs/boards/cc/apptools/<key>.html, run inside the app */
const TOOLS = [['hpwh', 'HPWH sizer', 'Heat pump water heater sizing + equipment'], ['pipe', 'Pipe pricer', 'Pipe material + labor'], ['plumb', 'Plumbing pricer', 'Plumbing scope pricing'],
  ['duct', 'Duct pricer', 'Duct takeoff pricing'], ['electrical', 'Electrical pricer', 'Electrical sub pricing'], ['permits', 'Permit Desk', 'Permits, portals, inspections'], ['proposal', 'Bid Board', 'Full Proposal Engine board']];
V.sales_kit = async () => `<h2>Tools</h2>` + TOOLS.filter(t => !['permits', 'proposal'].includes(t[0])).map(([k, l, d]) =>
  `<div class="card" onclick="OMNI.tool('${k}')"><div class="row"><div><div class="t">${esc(l)}</div><div class="s">${esc(d)}</div></div><span class="btn pri">Open</span></div></div>`).join('')
  + `<div class="muted">Tools run right here from your private SharePoint copies. Saved work goes back to the same SharePoint files the desktop board uses.</div>`;
V.sales_bidboard = async () => { setTimeout(() => OMNI.tool('proposal'), 0); return '<div class="loading">Opening the Bid Board…</div>'; };
V.construction_permitdesk = async () => { setTimeout(() => OMNI.tool('permits'), 0); return '<div class="loading">Opening the Permit Desk…</div>'; };

const ROOT = 'Open Construction Projects/Claude outputs';
let toolFrame = null;
async function openTool(key) {
  const t = TOOLS.find(x => x[0] === key) || [key, key];
  let box = document.getElementById('toolbox');
  if (!box) { box = document.createElement('div'); box.id = 'toolbox'; document.body.appendChild(box); }
  box.innerHTML = `<div class="toolbar"><button class="btn" id="toolclose">✕ Close</button><b>${esc(t[1])}</b><span class="muted" id="toolstat">Loading…</span></div><div id="tooltabs" class="seg" style="padding:6px 10px;margin:0;overflow-x:auto;flex-wrap:nowrap"></div><iframe id="toolframe"></iframe>`;
  box.hidden = false; document.getElementById('toolclose').onclick = () => { box.hidden = true; box.innerHTML = ''; toolFrame = null; };
  try {
    const html = await data('tool.' + key, () => spText(ROOT + '/boards/cc/apptools/' + key + '.html'), 60);
    toolFrame = document.getElementById('toolframe');
    toolFrame.src = URL.createObjectURL(new Blob([String(html)], {type: 'text/html'}));
    document.getElementById('toolstat').textContent = '';
  } catch (e) { document.getElementById('toolstat').textContent = e.message; }
}
/* SharePoint bridge for tool pages (same protocol as the desktop Command Center: cc-rpc read/list/readUri/write/copy) */
const WRITE_SCOPES = ['Files.ReadWrite.All'];
async function writeToken() {
  try { return (await pca.acquireTokenSilent({scopes: WRITE_SCOPES, account})).accessToken; }
  catch (e) { throw new Error('Saving needs one more Microsoft approval (Files.ReadWrite.All). Ask your Microsoft admin to approve it, then try again.'); }
}
function listingLine(x) { return `${x.name} (${x.folder ? 'folder' : 'file'}, ${x.size || 0} bytes) file:///${C.driveId}/${x.id}`; }
async function rpcOp(op, a) {
  if (op === 'read') { try { return await spText(ROOT + '/' + a.path); } catch (e) { if (/404|itemNotFound/i.test(e.message)) return null; throw e; } }
  if (op === 'list') { try { const r = await g(spItem(ROOT + (a.path ? '/' + a.path : '')) + '/children?$top=999&$select=name,id,size,folder,file');
      return (r.value || []).map(x => ({name: x.name, type: x.folder ? 'folder' : 'file', size: x.size || 0, uri: `file:///${C.driveId}/${x.id}`, id: x.id})); }
    catch (e) { if (/404|itemNotFound/i.test(e.message)) return null; throw e; } }
  if (op === 'readUri') {
    const m = String(a.uri).match(/^file:\/\/\/([^/]+)\/(.+)$/); if (!m) throw new Error('bad uri');
    const rest = decodeURIComponent(m[2]); const base = rest.includes('/') ? `/drives/${m[1]}/root:/${enc(rest)}:` : `/drives/${m[1]}/items/${rest}`;
    const it = await g(base + '?$select=name,id,size,folder,file');
    if (it.folder) { const r = await g(base + '/children?$top=999&$select=name,id,size,folder,file'); return (r.value || []).map(listingLine).join('\n'); }
    const r = await g(base + '/content', {raw: true}); return DEMO ? r : r.text();
  }
  if (op === 'write') {
    const tok = await writeToken();
    const r = await fetch(GRAPH + spItem(ROOT + '/' + a.path) + '/content', {method: 'PUT', headers: {Authorization: 'Bearer ' + tok, 'Content-Type': 'text/plain'}, body: a.content});
    if (!r.ok) throw new Error('Save failed: ' + r.status); Object.keys(mem).forEach(k => delete mem[k]); return true;
  }
  if (op === 'copy') {
    const tok = await writeToken();
    const r = await fetch(`${GRAPH}/drives/${C.driveId}/items/${a.itemId}/copy`, {method: 'POST', headers: {Authorization: 'Bearer ' + tok, 'Content-Type': 'application/json'}, body: JSON.stringify({parentReference: {driveId: C.driveId, id: a.parentId}, name: a.newName})});
    if (!r.ok && r.status !== 202) throw new Error('Copy failed: ' + r.status); return 'copied';
  }
  throw new Error('unknown op ' + op);
}
window.addEventListener('message', e => {
  const d = e.data || {}; if (d.type !== 'cc-tabs' || !toolFrame || e.source !== toolFrame.contentWindow) return;
  const bar = document.getElementById('tooltabs'); if (!bar) return; const curTab = bar.dataset.cur || (d.tabs[0] && d.tabs[0].key);
  bar.innerHTML = (d.tabs || []).map(t => `<button data-k="${esc(t.key)}" class="${t.key === curTab ? 'on' : ''}" style="white-space:nowrap">${esc(t.label)}${t.n ? ' (' + t.n + ')' : ''}</button>`).join('');
  bar.querySelectorAll('button').forEach(b => b.onclick = () => { bar.dataset.cur = b.dataset.k; bar.querySelectorAll('button').forEach(x => x.classList.toggle('on', x === b)); toolFrame.contentWindow.postMessage({type: 'cc-tab', tab: b.dataset.k}, location.origin); });
});
window.addEventListener('message', async e => {
  const d = e.data || {}; if (d.type !== 'cc-rpc' || !toolFrame || e.source !== toolFrame.contentWindow) return;
  let msg; try { msg = {type: 'cc-rpc-res', id: d.id, ok: true, result: await rpcOp(d.op, d.args || {})}; }
  catch (err) { msg = {type: 'cc-rpc-res', id: d.id, ok: false, error: {code: 'rpc_error', message: err.message}}; }
  try { e.source.postMessage(msg, location.origin); } catch (_) {}
});

V.finance_plan = async () => {
  const p = await L.plan(); const g0 = p.goal || {};
  const depts = p.depts || [], sellers = p.sellers || [];
  let h = `<div class="muted">2027 Business Plan targets (plan.json, read ${esc(p.source && p.source.readAt ? day(p.source.readAt) : '')})</div>
  <div class="kpis">${kpi(kmoney(g0.company), '2027 revenue goal', 'hot')}${kpi(kmoney(g0.install), 'Install goal')}${kpi(kmoney(g0.service), 'Service goal')}${kpi(pct((g0.netPct || 0) * 100), 'Net profit target')}</div>`;
  h += '<h2>Departments<small>2026 pace vs 2027 target</small></h2>' + depts.map(d => { const r = d.target ? d.pace2026 / d.target : 0;
    return `<div class="card"><div class="row"><div class="t">${esc(d.name)}</div><div class="amt">${kmoney(d.target)}</div></div><div class="s">2026 YTD ${kmoney(d.ytd2026)} · pace ${kmoney(d.pace2026)} = ${pct(r * 100)} of target · ${esc(d.field || '')} / ${esc(d.office || '')}</div><div class="bar"><i style="width:${Math.min(100, r * 100).toFixed(0)}%"></i></div></div>`; }).join('');
  h += '<h2>Sellers<small>2026 sold vs 2027 quota</small></h2>' + sellers.map(s => { const q = (s.install || 0) + (s.service || 0); const r = q ? s.sold2026 / q : 0;
    return `<div class="card"><div class="row"><div class="t">${esc(s.name)}</div><div class="amt">${kmoney(q)}</div></div><div class="s">2026 sold ${kmoney(s.sold2026)} · install ${kmoney(s.install)} · service ${kmoney(s.service)}</div><div class="bar"><i style="width:${Math.min(100, r * 100).toFixed(0)}%"></i></div></div>`; }).join('');
  if (p.pipeline) { const pl = p.pipeline; h += `<h2>Pipeline needed</h2><div class="kpis">${kpi(kmoney(pl.bidDollarsNeeded), 'Bid $ needed / yr')}${kpi(kmoney(pl.bidPerMonth), 'Bid $ / month')}${kpi(pct((pl.hitRate || 0) * 100), 'Hit rate (plan)')}${kpi(pl.membershipsPerMonth || '–', 'Memberships / month')}</div>`; }
  if (p.kpis) h += '<h2>KPIs</h2><div class="tbl"><table><tr><th>KPI</th><th>Now</th><th>Target</th><th>Owner</th></tr>' + p.kpis.map(k => `<tr><td>${esc(k[0])}</td><td>${esc(k[1])}</td><td>${esc(k[2])}</td><td>${esc(k[3])}</td></tr>`).join('') + '</table></div>';
  if (p.monthly) h += '<h2>Monthly ramp</h2><div class="tbl"><table><tr><th>Month</th><th class="n">Service</th><th class="n">Install</th></tr>' + p.monthly.map(m => `<tr><td>${esc(m[0])}</td><td class="n">${kmoney(m[1])}</td><td class="n">${kmoney(m[2])}</td></tr>`).join('') + '</table></div>';
  if (p.source && p.source.businessPlan) h += `<div class="acts"><a class="btn" href="${esc(p.source.businessPlan)}" target="_blank" rel="noopener">Business Plan</a><a class="btn" href="${esc(p.source.revenuePlan)}" target="_blank" rel="noopener">Revenue Plan</a></div>`;
  return h;
};

V.marketing_standout = async () => {
  const d = await L.ref(); const so = (d.docs && d.docs.standOut && d.docs.standOut.standOut) || [];
  return '<div class="card">' + so.map(s => `${s.l <= 1 ? `<h2>${esc(s.h)}</h2>` : `<div class="t" style="margin-top:12px">${esc(s.h)}</div>`}${(s.p || []).map(p => `<div class="p">${esc(p)}</div>`).join('')}`).join('') + '</div>';
};

/* ---------------- Command extras: queue requests, Drafts, Work, Systems Review, Dashboards ---------------- */
const MAIL_SCOPES = ['Mail.ReadWrite', 'Mail.Send'];
async function scopedToken(scopes, what) {
  try { return (await pca.acquireTokenSilent({scopes, account})).accessToken; }
  catch (e) { throw new Error(what + ' needs one more Microsoft approval. Ask your Microsoft admin to approve it, then try again.'); }
}
async function gw(path, opt, scopes, what) { /* Graph call with a write scope */
  if (DEMO) return {};
  const tok = await scopedToken(scopes, what); const h = {Authorization: 'Bearer ' + tok};
  let body = opt.body; if (body && typeof body !== 'string') { h['Content-Type'] = 'application/json'; body = JSON.stringify(body); } else if (body) h['Content-Type'] = 'text/plain';
  const r = await fetch(GRAPH + path, {method: opt.method || 'POST', headers: h, body});
  if (!r.ok && r.status !== 202) { let m = r.status + ''; try { m += ' ' + ((await r.json()).error.message || ''); } catch (_) {} throw new Error(m); }
  try { return await r.json(); } catch (_) { return {}; }
}
const PROJECTS = [['Org Instructions', '01a123f3-23f0-74f4-ba0d-6982b034e16d'], ['Proposal Engine', ''], ['General HVAC or Other Business Advice', ''], ['Reports & KPIs', ''], ['AI Branding/Online Presence Optimization', ''], ['Construction Weekly Wrap-Up', ''], ['BuildOps Entry & Admin', ''], ['Personal Command', '']];
async function queueWrite(fn) { /* read-modify-write boards/cc/inbox.json */
  const path = C.paths.boards + '/cc/inbox.json';
  let j; try { j = safeJSON(await spText(path)); } catch (_) { j = {docs: {}}; }
  j.docs = j.docs || {}; fn(j.docs); j.updatedAt = new Date().toISOString();
  await gw(spItem(path) + '/content', {method: 'PUT', body: JSON.stringify(j)}, WRITE_SCOPES, 'Saving');
  delete mem.queue;
}
async function queueAdd(item) {
  const id = 'i' + Date.now() + Math.random().toString(36).slice(2, 6); const now = new Date().toISOString();
  await queueWrite(d => { d[id] = Object.assign({kind: 'ask', ref: '', title: '', text: '', status: 'new', createdAt: now, _at: now, from: 'phone app', seat: SEAT.primary}, item); });
  return id;
}
V.command_queue = async () => {
  const q = await L.queue(); const items = Object.entries(q.docs || {}).map(([id, v]) => Object.assign({id}, v)).filter(i => !i._deleted);
  items.sort((a, b) => String(b.createdAt || '').localeCompare(String(a.createdAt || '')));
  const st = s => /done/i.test(s) ? 'g' : /error|fail/i.test(s) ? 'r' : /need|working|new/i.test(s) ? 'y' : '';
  const needs = items.filter(i => /needs you/i.test(i.status || ''));
  const card = i => `<div class="card ${/needs you/i.test(i.status) ? 'flag' : ''}"><div class="row"><div class="t">${esc(i.title || i.text || i.id)}</div><span class="pill ${st(i.status || '')}">${esc(i.status || '')}</span></div>
    <div class="s">${esc(when(i.createdAt))}${i.project ? ' · ' + esc(i.project) : ''}${i.kind ? ' · ' + esc(i.kind) : ''}</div>${i.text && i.title ? `<div class="p muted">${esc(String(i.text).slice(0, 300))}</div>` : ''}${i.reply ? `<div class="p">Claude: ${esc(String(i.reply).slice(0, 600))}</div>` : ''}
    ${/needs you/i.test(i.status) ? `<div class="addrow"><input id="ans_${i.id}" placeholder="Your answer"><button class="btn pri" onclick="OMNI.answer('${i.id}')">Send</button></div>` : ''}
    ${i.chat ? `<div class="acts"><a class="btn" href="https://claude.ai/chat/${esc(i.chat)}" target="_blank" rel="noopener">Open chat</a></div>` : ''}</div>`;
  return `<h2>Ask Claude<small>goes to the Command Center queue</small></h2><div class="card"><textarea id="qtext" rows="3" placeholder="What do you need done?"></textarea>
    <div class="addrow"><select id="qproj">${PROJECTS.map(p => `<option>${esc(p[0])}</option>`).join('')}</select><button class="btn pri" onclick="OMNI.ask()">Queue it</button></div>
    <div class="muted">The queue runner picks it up on its next run and answers here.</div></div>`
    + (needs.length ? '<h2>Needs you</h2>' + needs.map(card).join('') : '') + '<h2>Recent</h2>' + (items.filter(i => !needs.includes(i)).slice(0, 40).map(card).join('') || '<div class="empty">Queue is empty.</div>');
};
V.command_drafts = async () => {
  const r = await data('drafts', () => g(`/me/mailFolders/drafts/messages?$top=40&$orderby=lastModifiedDateTime desc&$select=subject,toRecipients,lastModifiedDateTime,bodyPreview,webLink`), 2);
  const ms = r.value || [];
  return `<h2>Drafts<small>${ms.length} in Outlook Drafts</small></h2>` + (ms.map(m => `<div class="card"><div class="row"><div class="t">${esc(m.subject || '(no subject)')}</div><div class="s">${esc(when(m.lastModifiedDateTime))}</div></div>
    <div class="s">To: ${esc((m.toRecipients || []).map(t => t.emailAddress.name || t.emailAddress.address).join(', ') || '–')}</div><div class="p muted">${esc((m.bodyPreview || '').slice(0, 220))}</div>
    <div class="acts"><button class="btn pri" onclick="OMNI.sendDraft('${m.id}', this)">Send</button><button class="btn" onclick="OMNI.mail('${m.id}')">Read / edit</button><a class="btn" href="${esc(m.webLink)}" target="_blank" rel="noopener">Outlook</a></div></div>`).join('') || '<div class="empty">No drafts.</div>');
};
V.command_work = async () => {
  const d = await data('work', () => spJSON(C.paths.boards + '/cc/work.json'), 10);
  const items = Object.entries(d.docs || {}).map(([id, v]) => Object.assign({id}, v)).filter(v => !v._deleted && !/archived|done/i.test(v.s || '') && !/nothing( else)? (is )?open|^done|closed out|no action|archive unless|\bdone\b[^.]*\.\s*archive/i.test(v.next || ''));
  items.forEach(v => { if (!v.n && /^[0-9a-f-]{36}$/.test(v.id)) v.chat = v.chat || v.id; });
  return `<h2>Work + next steps<small>${items.length} open</small></h2>` + (items.map(v => `<div class="card"><div class="t">${esc(v.n || v.title || (v.next || '').split(/[.;:]/)[0].slice(0, 70) || v.id)}</div><div class="s">${esc(v.p || v.proj || '')}${v.s ? ' · ' + esc(v.s) : ''}</div>${v.next ? `<div class="p">${esc(v.next)}</div>` : ''}
    ${v.chat || v.chat2 ? `<div class="acts"><a class="btn" href="https://claude.ai/chat/${esc(v.chat2 || v.chat)}" target="_blank" rel="noopener">Open chat</a></div>` : ''}</div>`).join('') || '<div class="empty">Nothing open.</div>');
};
V.command_systems = async () => {
  const d = await data('findings', () => spJSON(C.paths.boards + '/content/findings.json'), 10);
  const all = Object.entries(d.docs || {}).map(([id, v]) => Object.assign({id}, v)).filter(v => !v._deleted);
  const order = {new: 0, approved: 1, 'in-progress': 2, done: 3, skipped: 4};
  all.sort((a, b) => (order[a.status] ?? 5) - (order[b.status] ?? 5));
  const card = f => `<div class="card ${f.status === 'new' ? 'flag' : f.status === 'done' ? 'good' : ''}"><div class="row"><div class="t">${esc(f.title)}</div><span class="pill">${esc(f.status)}</span></div>
    <div class="s">${esc(f.area || '')} · ${esc(f.priority || '')} · ${esc(f.effort || '')}</div><div class="p">${esc(f.why || '')}</div>
    <div class="s">Now: ${esc(f.currentHome || '')}<br>Better: ${esc(f.recommendedHome || '')}</div>${f.needsTom ? `<div class="p muted">Needs you: ${esc(f.needsTom)}</div>` : ''}
    ${f.status === 'new' ? `<div class="acts"><button class="btn pri" onclick="OMNI.decide('${f.id}','approved')">Approve</button><button class="btn" onclick="OMNI.decide('${f.id}','skipped')">Skip</button></div>` : ''}</div>`;
  return `<div class="kpis">${kpi(all.filter(f => f.status === 'new').length, 'Waiting on you', 'hot')}${kpi(all.filter(f => /approved|in-progress/.test(f.status)).length, 'Approved / in progress')}${kpi(all.filter(f => f.status === 'done').length, 'Done')}${kpi(all.length, 'Cards')}</div>`
    + all.filter(f => !/done|skipped/.test(f.status)).map(card).join('') + '<h2>Done / skipped</h2>' + all.filter(f => /done|skipped/.test(f.status)).slice(0, 30).map(card).join('');
};
const DASHES = {
  'Construction': [['Job Board', 'UREWSNqroZwEJbZ3bQengX', 'Weekly job board, KPIs, Review tab'], ['Omni Lead-Time Log (old doc)', '9byfp1ihp7yxbHtoYkZuTm', 'Frozen 10/9; live log is Construction > Lead Times'], ['The Monroe Hub', 'GvQsbD6iJ3iuJ6SdHvNZBY', 'Monroe job hub'], ['Construction Wrap-Up (page)', 'AqGuDhSzWcWfodSUXTNTSh', 'Older wrap-up page']],
  'Sales & Bidding': [['Proposal Engine', '1KjFibJnA9wChDTxSX54yH', 'Bid board, CRM queue, filing, catalog, rebates'], ['Plumbing Pricer', 'F2yNKJeEM4fGZTYRMB97YK', 'Sub price range'], ['Duct Pricer', 'R5LuNWXNvoAP7k975G6nhC', 'QuickTin-based duct cost'], ['Electrical Pricer', 'BzzNWWVARUM3rtn9D1JLTm', 'Apex-based electrical cost']],
  'Finance': [['Omni Department Scorecard', 'Ufoo4dhbbBEGJ3VsNspyuv', 'Department metrics'], ['2027 Revenue Plan', '5CBBdBHQT3muxTCpXd5dho', 'Plan sheet'], ['2027 Business Plan', '772hxWaJn6mjoW3CaFSzEV', 'Plan doc']],
  'Other': [['Desktop Command Center', 'Ni1J4DsD6J2it5XdBC1Epi', 'Full board on claude.ai'], ['Uniform Order', 'DmjEGWzYQApfNVmiU3Tfws', 'Uniform ordering'], ['BuildOps Entry Desk', 'NHGSPaFamtjWMqRdbTVNMf', 'Job/project setup queue'], ['QBO Change Log', 'VdFfd5pLWTSE6Zhc9xB7X7', 'QuickBooks change batches']]
};
V.command_dash = async () => Object.entries(DASHES).map(([g0, arr]) => `<h2>${esc(g0)}</h2>` + arr.map(([n, id, d]) =>
  `<a class="card" style="display:block;text-decoration:none;color:inherit" href="https://claude.ai/artifact/${n === 'Desktop Command Center' ? (SEAT_BOARDS[SEAT.primary] || id) : id}" target="_blank" rel="noopener"><div class="row"><div><div class="t">${esc(n)}${n === 'Desktop Command Center' ? ' (' + esc(SEAT.primary) + ')' : ''}</div><div class="s">${esc(d)}</div></div><span class="btn">Open ↗</span></div></a>`).join('')).join('')
  + `<div class="muted">These open in the Claude app.</div>`;

/* ---------------- Seats: which Claude seat is primary (SharePoint boards/cc/active.json, flipped with ledger line 1) ---------------- */
const SEAT_BOARDS = {'tcm@': 'JemzpN7zAm6YNY5xhoV6YK', 'engineering@': 'Ni1J4DsD6J2it5XdBC1Epi'};
L.seat = () => data('seat', () => spJSON(C.paths.boards + '/cc/active.json'), 2);
let SEAT = store.get('seat', {primary: 'tcm@'});
async function refreshSeat() {
  try { const s = await L.seat(); if (s && s.primary) { SEAT = s; store.set('seat', s); } } catch (_) {}
  const b = document.getElementById('seatbadge');
  if (b) { const eng = /engineering/.test(SEAT.primary); b.textContent = eng ? 'ENG@' : 'TCM@'; b.className = 'seatbadge' + (eng ? ' eng' : ''); b.title = 'Active Claude seat: ' + SEAT.primary; }
  C.links.claudeBoard = 'https://claude.ai/artifact/' + (SEAT_BOARDS[SEAT.primary] || SEAT_BOARDS['tcm@']);
  return SEAT;
}
V.command_seats = async () => {
  const s = await refreshSeat(); const eng = /engineering/.test(s.primary);
  const q = await L.queue().catch(() => ({docs: {}}));
  const pending = Object.values(q.docs || {}).filter(i => i.kind === 'seat' && /new|working/.test(i.status || ''));
  return `<div class="kpis">${kpi(esc(s.primary), 'Active Claude seat', 'hot')}${kpi(esc(s.since ? when(s.since) : '–'), 'Since')}</div>
  <div class="card ${eng ? 'flag' : 'good'}"><div class="t">${eng ? 'Takeover: engineering@ is running things' : 'Normal: tcm@ is primary'}</div>
  <div class="s">The phone app reads the same SharePoint data either way. The seat decides which Claude account runs the queue, tasks and the desktop board.</div>
  <div class="acts"><a class="btn pri" href="${C.links.claudeBoard}" target="_blank" rel="noopener">Open ${esc(s.primary)} desktop board</a></div></div>
  <h2>Switch seats</h2><div class="card"><div class="s">Queues the switch for Claude. It flips the ledger, the mirror log and active.json, and turns the mirror tasks on or off (omni-handoff-ledger). The app follows on the next refresh.</div>
  <div class="acts">${eng ? `<button class="btn pri" onclick="OMNI.seat('hand back')">Hand back to tcm@</button>` : `<button class="btn pri" onclick="OMNI.seat('take over')">Take over with engineering@</button>`}</div>
  ${pending.length ? `<div class="p muted">Switch already queued ${esc(when(pending[0].createdAt))}: "${esc(pending[0].text)}"</div>` : ''}</div>`;
};

/* ---------------- actions ---------------- */
/* ---------- Uniforms: SharePoint lists "Uniform Order Lines" (current quarter) + "Uniform Order Lines Archive" on /sites/OmniMech ---------- */
const UNI = {
  site: 'omnimechusa.sharepoint.com:/sites/OmniMech', live: 'Uniform Order Lines', arch: 'Uniform Order Lines Archive',
  emps: ['Jack Tacchi', 'Josh Beggs', 'Paul Lutsch', 'Kaydin Johnson', 'Jeffrey Randall', 'Quentin Harding', 'Roland Gobel', 'Christopher Hook', 'Dillon Willhoite', 'Brian Johnson', 'Parker Rowand', 'Thomas Mitchell', 'Lucas Barker', 'Tim', 'Darian'],
  sizes: ['S', 'M', 'L', 'XL', '2XL', '3XL', '4XL'],
  field: d => d === 'Service' || d === 'Install',
  meta(k, d) { const f = UNI.field(d);
    return ({
      tee: f ? ['Tee Shirt', 'Hi-Vis', '8300 Gildan DryBlend Pocket Tee'] : ['Tee Shirt', 'Black', 'Gildan 2000 Short Sleeve Tee'],
      longSleeve: f ? ['Long Sleeve', 'Hi-Vis', 'Hi-Vis Long Sleeve Tee'] : ['Long Sleeve', 'Black', 'Gildan 2000 Long Sleeve Tee'],
      sweatshirt: ['Sweatshirt', 'Black', '18500 Gildan Heavy Blend Hoodie' + (f ? ', Field Art' : ', Office Art')],
      jacket: ['Jacket', 'Black', 'Carhartt Storm Defender Shoreline Jacket'],
      hiVisJacket: ['Hi-Vis Jacket', 'Safety Green', 'Port Authority Enhanced Visibility Challenger Jacket'],
      ballCap: ['Ball Cap', 'Charcoal / Kelly Green', '112 Trucker Cap'],
      beanie: ['Beanie', 'Black', 'Omni Beanie'],
      teeHiVis: ['Tee Shirt', 'Hi-Vis', '8300 Gildan DryBlend Pocket Tee, office allowance'],
      lsHiVis: ['Long Sleeve', 'Hi-Vis', 'Hi-Vis Long Sleeve Tee, office allowance']})[k]; },
  keys: d => ['tee', 'longSleeve', 'sweatshirt', 'jacket'].concat(UNI.field(d) ? ['hiVisJacket'] : []).concat(['ballCap', 'beanie']).concat(d === 'Office' ? ['teeHiVis', 'lsHiVis'] : []),
  sized: k => !/ballCap|beanie/.test(k), long: k => /tee|longSleeve|sweatshirt|teeHiVis|lsHiVis/.test(k),
  caps: [{label: 'jackets', max: 2, test: l => /jacket/i.test(l.GarmentType)}, {label: 'office hi-vis tees', max: 2, test: l => l.Department === 'Office' && /tee shirt/i.test(l.GarmentType) && /hi-vis/i.test(l.Color)}, {label: 'office hi-vis long sleeves', max: 2, test: l => l.Department === 'Office' && /long sleeve/i.test(l.GarmentType) && /hi-vis/i.test(l.Color)}],
  cart: []
};
L.uniSite = () => data('uniSite', () => g(`/sites/${UNI.site}?$select=id`), 1440);
L.uniLines = () => data('uniLines', async () => {
  const s = await L.uniSite(); const get = async t => { let u = `/sites/${s.id}/lists/${encodeURIComponent(t)}/items?$expand=fields&$top=500`, out = [];
    while (u) { const r = await g(u.replace(GRAPH, '')); out = out.concat((r.value || []).map(x => Object.assign({_id: x.id, _list: t}, x.fields))); u = r['@odata.nextLink'] || ''; } return out; };
  const [live, arch] = await Promise.all([get(UNI.live), get(UNI.arch).catch(() => [])]);
  return {live, arch};
}, 15);
const uniYear = l => String(l.OrderDate || l.Created || '').slice(0, 4);
V.marketing_uniforms = async () => {
  const sub = store.get('uniView', 'orders');
  const seg = `<div class="seg">${[['orders', 'This quarter'], ['buy', 'Buy list'], ['people', 'People'], ['new', 'New order'], ['reports', 'Reports']].map(([k, n]) => `<button class="${sub === k ? 'on' : ''}" onclick="OMNI.uniView('${k}')">${n}</button>`).join('')}</div>`;
  let d; try { d = await L.uniLines(); } catch (e) { return seg + `<div class="card bad"><div class="t">Can't read the uniform lists</div><div class="s">${esc(e.message)}</div></div>`; }
  const live = d.live.filter(l => !/test/i.test(l.EmployeeName || '')), all = live.concat(d.arch), yr = String(new Date().getFullYear());
  const qty = a => a.reduce((s, l) => s + (num(l.Quantity) || 1), 0);
  let h = `<div class="kpis">${kpi(qty(live), 'Pieces this quarter', live.length ? 'hot' : '')}${kpi(new Set(live.map(l => l.EmployeeName)).size, 'People ordering')}${kpi(qty(all.filter(l => uniYear(l) === yr)), 'Pieces this year')}${kpi(d.arch.length, 'Archived lines')}</div>` + seg;
  if (sub === 'orders') {
    const by = {}; live.forEach(l => (by[l.EmployeeName] = by[l.EmployeeName] || []).push(l));
    h += Object.keys(by).sort().map(n => `<div class="card"><div class="row"><div class="t">${esc(n)}</div><div class="s">${esc(by[n][0].Department || '')}</div></div>${by[n].map(l => `<div class="s">${esc(l.Quantity)} × ${esc(l.GarmentType)} · ${esc(l.Color)} · ${esc(l.Size)}${l.Length && l.Length !== 'Standard' ? ' ' + esc(l.Length) : ''} <span class="muted">${esc(day(l.OrderDate || l.Created))}</span></div>`).join('')}</div>`).join('') || '<div class="empty">No orders this quarter.</div>';
    if (d.live.length !== live.length) h += `<div class="muted">${d.live.length - live.length} test line(s) hidden.</div>`;
  } else if (sub === 'buy') {
    const t = {}; live.forEach(l => { const k = [l.GarmentLabel || l.GarmentType, l.Size, l.Length || 'Standard'].join('|'); t[k] = (t[k] || 0) + (num(l.Quantity) || 1); });
    const rows = Object.entries(t).sort();
    h += `<div class="acts"><button class="btn pri" onclick="OMNI.uniExport()">Download Excel</button></div><h2>What to order from the vendor</h2>` + (rows.map(([k, q]) => { const [g1, s, len] = k.split('|'); return `<div class="card"><div class="row"><div class="t">${esc(g1)}</div><div class="amt">${q}</div></div><div class="s">Size ${esc(s)}${len !== 'Standard' ? ' · ' + esc(len) : ''}</div></div>`; }).join('') || '<div class="empty">Nothing to order.</div>');
  } else if (sub === 'people') {
    const ppl = [...new Set(all.filter(l => uniYear(l) === yr).map(l => l.EmployeeName))].filter(n => n && !/test/i.test(n)).sort();
    h += '<h2>Annual allowances (' + yr + ')</h2>' + ppl.map(n => { const mine = all.filter(l => l.EmployeeName === n && uniYear(l) === yr);
      return `<div class="card"><div class="row"><div class="t">${esc(n)}</div><div class="amt">${qty(mine)} pcs</div></div><div class="s">${UNI.caps.map(c => { const u = qty(mine.filter(c.test)); return u ? `${u}/${c.max} ${c.label}${u >= c.max ? ' (maxed)' : ''}` : ''; }).filter(Boolean).join(' · ') || 'No capped items yet'}</div></div>`; }).join('') || '<div class="empty">No orders this year.</div>';
  } else if (sub === 'new') {
    const f = store.get('uniForm', {name: '', dept: 'Install', key: 'tee', size: 'L', len: 'Standard', qty: 1});
    const opt = (a, v) => a.map(x => `<option ${x === v ? 'selected' : ''}>${esc(x)}</option>`).join('');
    const keys = UNI.keys(f.dept); if (!keys.includes(f.key)) f.key = keys[0];
    h += `<div class="card"><div class="t">Order for an employee</div><div class="s">Writes straight to the "Uniform Order Lines" list, so the quarterly report picks it up.</div>
    <div class="form"><label>Name<select id="uName" onchange="OMNI.uniForm('name',this.value)"><option value="">Pick a name</option>${opt(UNI.emps, f.name)}</select></label>
    <label>Department<select onchange="OMNI.uniForm('dept',this.value)">${opt(['Office', 'Service', 'Install'], f.dept)}</select></label>
    <label>Garment<select onchange="OMNI.uniForm('key',this.value)">${keys.map(k => `<option value="${k}" ${k === f.key ? 'selected' : ''}>${esc(UNI.meta(k, f.dept)[0])} (${esc(UNI.meta(k, f.dept)[1])})</option>`).join('')}</select></label>
    ${UNI.sized(f.key) ? `<label>Size<select onchange="OMNI.uniForm('size',this.value)">${opt(UNI.sizes, f.size)}</select></label>` : ''}
    ${UNI.long(f.key) ? `<label>Length<select onchange="OMNI.uniForm('len',this.value)">${opt(['Standard', 'Tall'], f.len)}</select></label>` : ''}
    <label>Quantity<select onchange="OMNI.uniForm('qty',+this.value)">${opt(['1', '2'], String(f.qty))}</select></label></div>
    <div class="acts"><button class="btn" onclick="OMNI.uniAdd()">Add to order</button></div></div>`;
    if (UNI.cart.length) h += `<h2>This order</h2>` + UNI.cart.map((c, i) => `<div class="card"><div class="row"><div class="t">${esc(c.qty)} × ${esc(c.type)} · ${esc(c.color)} · ${esc(c.size)}</div><button class="btn" onclick="OMNI.uniDrop(${i})">Remove</button></div><div class="s">${esc(c.name)} · ${esc(c.dept)}</div></div>`).join('') + `<div class="acts"><button class="btn pri" onclick="OMNI.uniSubmit()">Submit ${UNI.cart.length} line(s)</button></div>`;
  } else {
    const reps = await L.uniformReports().catch(() => ({})); const rs = reps.value || [];
    h += `<div class="card"><div class="s">The "Quarterly Uniform Order Report" Power Automate flow emails admin@ each quarter from these same lists.</div></div>` + (rs.slice(0, 8).map(m => `<div class="card"><div class="row"><div class="t">${esc(m.subject)}</div><div class="s">${esc(day(m.receivedDateTime))}</div></div><div class="acts">${m.hasAttachments ? `<button class="btn pri" onclick="OMNI.report('${m.id}')">Open report</button>` : ''}<a class="btn" href="${esc(m.webLink)}" target="_blank" rel="noopener">Email</a></div></div>`).join('') || '<div class="empty">No uniform report emails in your mailbox.</div>');
  }
  return h;
};

/* ---------- Lead-Time Log: SharePoint boards/leadtime/items.json (team edits) + pos.json (PC script, open BuildOps POs) ---------- */
L.lead = () => data('lead', () => spJSON(board('leadtime/items.json')), 5);
L.leadPOs = () => data('leadPOs', () => spJSON(board('leadtime/pos.json')), 60);
const LEAD_ST = ['NOT ORDERED', 'ORDERED', 'WAITING', 'CHECK', 'DELIVERED'];
async function leadWrite(fn) {
  const raw = await rpcOp('read', {path: 'boards/leadtime/items.json'}); const j = raw ? JSON.parse(raw) : {docs: {}};
  const now = new Date().toISOString(); fn(j.docs, now); j.updatedAt = now;
  await rpcOp('write', {path: 'boards/leadtime/items.json', content: JSON.stringify(j, null, 1)}); delete mem.lead;
}
V.construction_leadtime = async () => {
  const f = store.get('leadView', 'open');
  let j = {docs: {}}; try { j = await L.lead(); } catch (_) {}
  const it = Object.entries(j.docs || {}).filter(([, x]) => !x._deleted).map(([id, x]) => Object.assign({id}, x));
  const open = it.filter(x => x.status !== 'DELIVERED'), t = todayYMD();
  const late = open.filter(x => x.eta && x.eta < t);
  const cnt = s => open.filter(x => x.status === s).length;
  let h = `<div class="kpis">${kpi(cnt('NOT ORDERED'), 'Not ordered', cnt('NOT ORDERED') ? 'bad' : '')}${kpi(cnt('WAITING') + cnt('ORDERED'), 'On order')}${kpi(cnt('CHECK'), 'Check', cnt('CHECK') ? 'hot' : '')}${kpi(late.length, 'Past ETA', late.length ? 'bad' : '')}</div>`;
  h += `<div class="seg">${[['open', 'Open'], ['NOT ORDERED', 'Not ordered'], ['WAITING', 'On order'], ['CHECK', 'Check'], ['DELIVERED', 'Delivered'], ['pos', 'BuildOps POs'], ['add', '+ Add']].map(([k, n]) => `<button class="${f === k ? 'on' : ''}" onclick="OMNI.leadView('${k}')">${n}</button>`).join('')}</div>`;
  if (f === 'add') return h + `<div class="card"><div class="form"><label>Item<input id="lItem"></label><label>Job<input id="lJob"></label><label>Vendor<input id="lVen"></label><label>Lead time<input id="lLead" placeholder="e.g. 6-8 weeks"></label><label>ETA<input id="lEta" type="date"></label><label>Status<select id="lSt">${LEAD_ST.map(s => `<option>${s}</option>`).join('')}</select></label><label>Note<input id="lNote"></label></div><div class="acts"><button class="btn pri" onclick="OMNI.leadAdd()">Save</button></div></div>`;
  if (f === 'pos') { let p = null; try { p = await L.leadPOs(); } catch (_) {}
    if (!p) return h + `<div class="card"><div class="s">Open BuildOps POs land here once the PC job "Omni Lead Times" runs (C:\\OmniAgent\\leadtime, daily 6:10 AM, no Claude).</div></div>`;
    const tracked = new Set(it.map(x => x.po).filter(Boolean));
    return h + `<div class="muted">From BuildOps ${esc(when(p.updatedAt))} · ${p.items.length} open POs</div>` + p.items.map(x => `<div class="card ${tracked.has(x.po) ? '' : 'flag'}"><div class="row"><div class="t">PO ${esc(x.po)} · ${esc(x.vendor || '')}</div><div class="amt">${x.total ? money(x.total) : ''}</div></div><div class="s">${esc(x.job || '')} · ${esc(x.status || '')} · ${esc(x.date ? day(x.date + 'T12:00:00Z') : '')}${x.needBy ? ' · need by ' + esc(x.needBy) : ''}</div>${x.desc ? `<div class="p muted">${esc(x.desc)}</div>` : ''}${tracked.has(x.po) ? '' : `<div class="acts"><button class="btn" onclick="OMNI.leadTrack('${esc(x.po)}')">Track lead time</button></div>`}</div>`).join(''); }
  const show = (f === 'open' ? open : f === 'WAITING' ? it.filter(x => x.status === 'WAITING' || x.status === 'ORDERED') : it.filter(x => x.status === f)).sort((a, b) => (a.eta || '9').localeCompare(b.eta || '9'));
  h += show.map(x => { const cls = x.status === 'DELIVERED' ? '' : x.eta && x.eta < t ? 'bad' : x.status === 'NOT ORDERED' ? 'bad' : x.status === 'CHECK' ? 'flag' : '';
    return `<div class="card ${cls}"><div class="row"><div class="t">${esc(x.item)}</div><div class="s">${esc(x.status)}</div></div><div class="s">${esc(x.job || '')}${x.vendor ? ' · ' + esc(x.vendor) : ''}${x.lead ? ' · ' + esc(x.lead) : ''}</div><div class="s">ETA ${x.eta ? esc(day(x.eta + 'T12:00:00Z')) : esc(x.etaText || 'none')}${x.po ? ' · PO ' + esc(x.po) : ''}</div>${x.note ? `<div class="p muted">${esc(x.note)}</div>` : ''}
    <div class="acts">${LEAD_ST.filter(s => s !== x.status).map(s => `<button class="btn" onclick="OMNI.leadSet('${x.id}','status','${s}')">${s === 'DELIVERED' ? 'Delivered' : s[0] + s.slice(1).toLowerCase()}</button>`).join('')}<input type="date" value="${esc(x.eta || '')}" onchange="OMNI.leadSet('${x.id}','eta',this.value)" style="max-width:150px"></div></div>`; }).join('') || '<div class="empty">Nothing here.</div>';
  return h;
};
const OMNI_X = {
  uniView(k) { store.set('uniView', k); render(); },
  uniForm(k, v) { const f = store.get('uniForm', {name: '', dept: 'Install', key: 'tee', size: 'L', len: 'Standard', qty: 1}); f[k] = v; store.set('uniForm', f); render(); },
  async uniAdd() {
    const f = store.get('uniForm', {name: '', dept: 'Install', key: 'tee', size: 'L', len: 'Standard', qty: 1}); if (!f.name) return toast('Pick a name first');
    const [type, color, product] = UNI.meta(f.key, f.dept); const line = {name: f.name, dept: f.dept, type, color, product, size: UNI.sized(f.key) ? f.size : 'OSFM', len: UNI.long(f.key) ? f.len : 'Standard', qty: +f.qty || 1};
    const d = await L.uniLines().catch(() => ({live: [], arch: []})); const yr = String(new Date().getFullYear());
    const lf = {EmployeeName: f.name, Department: f.dept, GarmentType: type, Color: color};
    for (const c of UNI.caps) if (c.test(lf)) { const used = d.live.concat(d.arch).filter(l => l.EmployeeName === f.name && uniYear(l) === yr && c.test(l)).reduce((s, l) => s + (num(l.Quantity) || 1), 0) + UNI.cart.filter(x => x.name === f.name && c.test({EmployeeName: x.name, Department: x.dept, GarmentType: x.type, Color: x.color})).reduce((s, x) => s + x.qty, 0);
      if (used + line.qty > c.max) return toast(`${f.name} has ${used} of ${c.max} ${c.label} this year`); }
    UNI.cart.push(line); render();
  },
  uniDrop(i) { UNI.cart.splice(i, 1); render(); },
  async uniSubmit() {
    try { const s = await L.uniSite(); const day0 = todayYMD() + 'T12:00:00Z';
      for (const c of UNI.cart) { const label = `${c.type} - ${c.color} (${c.product})`;
        await gw(`/sites/${s.id}/lists/${encodeURIComponent(UNI.live)}/items`, {body: {fields: {Title: `${c.name} - ${label}`, EmployeeName: c.name, Department: c.dept, GarmentType: c.type, GarmentLabel: label, Color: c.color, Size: c.size, Length: c.len, Quantity: c.qty, OrderDate: day0}}}, ['Sites.ReadWrite.All'], 'Placing uniform orders'); }
      toast(`Saved ${UNI.cart.length} line(s)`); UNI.cart = []; delete mem.uniLines; store.set('uniView', 'orders'); render();
    } catch (e) { toast(e.message); }
  },
  async uniExport() { await loadXlsx(); const d = await L.uniLines();
    const rows = d.live.map(l => ({Date: String(l.OrderDate || l.Created || '').slice(0, 10), Name: l.EmployeeName, Department: l.Department, Garment: l.GarmentLabel || l.GarmentType, Color: l.Color, Size: l.Size, Length: l.Length, Quantity: l.Quantity}));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Orders'); XLSX.writeFile(wb, `omni-uniform-orders-${todayYMD()}.xlsx`); },
  leadView(k) { store.set('leadView', k); render(); },
  async leadSet(id, k, v) { try { await leadWrite((d, now) => { if (d[id]) { d[id][k] = v; d[id]._at = now; d[id].by = 'phone'; } }); toast('Saved'); render(); } catch (e) { toast(e.message); } },
  async leadAdd() { const v = id => (document.getElementById(id) || {}).value || ''; if (!v('lItem')) return toast('Item is required');
    try { await leadWrite((d, now) => { d['lt-' + Date.now().toString(36)] = {item: v('lItem'), job: v('lJob'), vendor: v('lVen'), lead: v('lLead'), eta: v('lEta'), status: v('lSt'), note: v('lNote'), source: 'phone', addedAt: now, _at: now}; }); store.set('leadView', 'open'); toast('Added'); render(); } catch (e) { toast(e.message); } },
  async leadTrack(po) { const p = await L.leadPOs(); const x = (p.items || []).find(i => i.po === po); if (!x) return;
    try { await leadWrite((d, now) => { d['lt-po-' + po] = {item: x.desc || ('PO ' + po), job: x.job || '', vendor: x.vendor || '', po, eta: x.needBy || '', status: 'ORDERED', source: 'BuildOps PO', addedAt: now, _at: now}; }); toast('Tracking PO ' + po); render(); } catch (e) { toast(e.message); } }
};

/* ---------- Admin To Do = Microsoft List "Management To Do's" (Joyce's OneDrive list, installed as a Chrome app on Tom's PC) ---------- */
const MGMT = {site: 'omnimechusa-my.sharepoint.com:/personal/joyce_aguilar_omnimechusa_com', list: 'e2570a0f-edd4-449b-8e33-70c2a6c4c0fc',
  web: 'https://omnimechusa-my.sharepoint.com/personal/joyce_aguilar_omnimechusa_com/Lists/Management%20To%20Dos/AllItems.aspx'};
L.mgmtSite = () => data('mgmtSite', () => g(`/sites/${MGMT.site}?$select=id`), 1440);
L.mgmt = () => data('mgmt', async () => { const s = await L.mgmtSite();
  const r = await g(`/sites/${s.id}/lists/${MGMT.list}/items?$expand=fields&$top=500`);
  const items = (r.value || []).map(x => Object.assign({_id: x.id}, x.fields));
  const ids = [...new Set(items.map(x => x.Assignedto0LookupId).filter(Boolean))], names = {};
  await Promise.all(ids.map(async i => { try { const u = await g(`/sites/${s.id}/lists/User%20Information%20List/items/${i}?$expand=fields($select=Title)`); names[i] = u.fields.Title; } catch (_) {} }));
  items.forEach(x => x._who = names[x.Assignedto0LookupId] || ''); return items; }, 5);
V.command_admin = async () => {
  const f = store.get('mgmtView', 'open'); let items;
  try { items = await L.mgmt(); } catch (e) { return `<h2>Management To Do's</h2><div class="card bad"><div class="s">${esc(e.message)}</div><div class="acts"><a class="btn pri" href="${MGMT.web}" target="_blank" rel="noopener">Open the list</a></div></div>`; }
  const done = x => /completed|duplicate|by design|won't fix/i.test(x.Status || ''), t = todayYMD();
  const open = items.filter(x => !done(x)), od = open.filter(x => x.DueDate && x.DueDate.slice(0, 10) < t);
  const pr = {Critical: 0, High: 1, Normal: 2, Low: 3};
  let h = `<h2>Management To Do's<small>Joyce's list</small></h2><div class="kpis">${kpi(open.length, 'Open')}${kpi(od.length, 'Overdue', od.length ? 'bad' : '')}${kpi(open.filter(x => x.Priority === 'Critical').length, 'Critical', 'hot')}${kpi(open.filter(x => x.Status === 'Blocked').length, 'Blocked')}</div>`;
  h += `<div class="seg">${[['open', 'Open'], ['mine', 'Mine'], ['done', 'Done'], ['add', '+ Add']].map(([k, n]) => `<button class="${f === k ? 'on' : ''}" onclick="OMNI.mgmtView('${k}')">${n}</button>`).join('')}<a class="btn" style="margin-left:auto" href="${MGMT.web}" target="_blank" rel="noopener">Open list</a></div>`;
  if (f === 'add') return h + `<div class="card"><div class="form"><label>Task<input id="mT"></label><label>Priority<select id="mP"><option>Normal</option><option>High</option><option>Critical</option><option>Low</option></select></label><label>Due<input id="mD" type="date" value="${t}"></label><label>Description<input id="mN"></label></div><div class="acts"><button class="btn pri" onclick="OMNI.mgmtAdd()">Add</button></div></div>`;
  const me = (store.get('me', {}) || {}).displayName || 'Tom';
  const show = (f === 'done' ? items.filter(done) : f === 'mine' ? open.filter(x => /tom|thomas/i.test(x._who)) : open)
    .sort((a, b) => f === 'done' ? String(b.Modified).localeCompare(String(a.Modified)) : ((pr[a.Priority] ?? 2) - (pr[b.Priority] ?? 2)) || String(a.DueDate || '9').localeCompare(String(b.DueDate || '9'))).slice(0, f === 'done' ? 25 : 200);
  h += show.map(x => { const due = x.DueDate ? x.DueDate.slice(0, 10) : ''; const cls = done(x) ? '' : due && due < t ? 'bad' : x.Priority === 'Critical' ? 'flag' : '';
    return `<div class="card ${cls} todo">${done(x) ? '' : `<button class="check" aria-label="Done" onclick="OMNI.mgmtDone(this,'${x._id}')"></button>`}<div class="grow"><div class="t">${esc(x.Title)}</div><div class="s">${esc(x.Status || '')} · ${esc(x.Priority || 'Normal')}${x._who ? ' · ' + esc(x._who) : ''}${due ? ' · due ' + esc(day(due + 'T12:00:00Z')) : ''}</div>${x.Description ? `<div class="p muted">${esc(String(x.Description).replace(/<[^>]+>/g, ' ').trim().slice(0, 240))}</div>` : ''}</div></div>`; }).join('') || '<div class="empty">All clear.</div>';
  return h;
};
Object.assign(OMNI_X, {
  mgmtView(k) { store.set('mgmtView', k); render(); },
  async mgmtDone(el, id) { el.classList.add('on'); try { const s = await L.mgmtSite(); await gw(`/sites/${s.id}/lists/${MGMT.list}/items/${id}/fields`, {method: 'PATCH', body: {Status: 'Completed'}}, ['Sites.ReadWrite.All'], 'Checking off Management To Do\'s'); delete mem.mgmt; toast('Completed'); render(); } catch (e) { el.classList.remove('on'); toast(e.message); } },
  async mgmtAdd() { const v = id => (document.getElementById(id) || {}).value || ''; if (!v('mT')) return toast('Task is required');
    try { const s = await L.mgmtSite(); await gw(`/sites/${s.id}/lists/${MGMT.list}/items`, {body: {fields: Object.assign({Title: v('mT'), Priority: v('mP'), Status: 'New', Description: v('mN')}, v('mD') ? {DueDate: v('mD') + 'T12:00:00Z'} : {})}}, ['Sites.ReadWrite.All'], 'Adding to Management To Do\'s');
      delete mem.mgmt; store.set('mgmtView', 'open'); toast('Added'); render(); } catch (e) { toast(e.message); } }
});

window.OMNI = {
  go, tool: openTool,
  async seat(action) { try { await queueAdd({kind: 'seat', title: action === 'take over' ? 'Take over: make engineering@ primary' : 'Hand back: make tcm@ primary', text: action, project: 'Org Instructions', projectId: '01a123f3-23f0-74f4-ba0d-6982b034e16d'}); toast('Seat switch queued for Claude'); render(); } catch (e) { toast(e.message); } }, set(k, v) { store.set(k, v); render(); },
  pickList(k, id) { store.set(k, id); render(); },
  adminSet(type, v) { if (type === 'link' && !/^https?:\/\//.test(v)) return toast('Paste the full https:// link'); store.set('admin', type === 'todo' ? {type, id: v} : {type, url: v, name: 'Admin To Do'}); render(); },
  adminReset() { store.set('admin', null); render(); },
  async done(btn, list, id) { btn.classList.add('done'); try { await g(`/me/todo/lists/${list}/tasks/${id}`, {method: 'PATCH', body: {status: 'completed'}}); delete mem['tasks.' + list]; btn.closest('.card').style.opacity = .35; toast('Done ✓'); } catch (e) { btn.classList.remove('done'); toast(e.message); } },
  async addTask(list) { const i = document.getElementById('newTask'); const t = i.value.trim(); if (!t) return; i.value = '';
    try { await g(`/me/todo/lists/${list}/tasks`, {method: 'POST', body: {title: t, dueDateTime: {dateTime: todayYMD() + 'T17:00:00', timeZone: 'Pacific Standard Time'}}}); delete mem['tasks.' + list]; render(); toast('Added'); } catch (e) { toast(e.message); } },
  async mail(id) { sheet('<div class="loading">Opening…</div>');
    try { const m = await g(`/me/messages/${id}?$select=subject,from,receivedDateTime,body,webLink,toRecipients,isDraft,isRead,internetMessageId`);
      const draft = m.isDraft;
      sheet(`<div class="t" style="font-size:17px;padding-right:30px">${esc(m.subject)}</div><div class="s">${draft ? 'Draft to ' + esc((m.toRecipients || []).map(t => t.emailAddress.address).join(', ')) : esc(m.from && m.from.emailAddress.name)} · ${esc(dfmt(m.receivedDateTime, {month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit'}))}</div>
      <div class="acts">${draft ? `<button class="btn pri" onclick="OMNI.sendDraft('${id}', this)">Send draft</button>` : `<button class="btn pri" onclick="OMNI.compose('${id}','reply')">Reply</button><button class="btn" onclick="OMNI.compose('${id}','replyAll')">Reply all</button><button class="btn" onclick="OMNI.mailAct('${id}','archive', this)">Archive</button><button class="btn" onclick="OMNI.mailAct('${id}','${m.isRead ? 'unread' : 'read'}', this)">Mark ${m.isRead ? 'unread' : 'read'}</button>`}<a class="btn" href="${esc(m.webLink)}" target="_blank" rel="noopener">Outlook</a></div>
      <div id="composer"></div>
      ${draft ? '' : `<div class="addrow"><input id="askc" placeholder="Ask Claude about this email (draft a reply, file it, make a To Do…)"><button class="btn" onclick="OMNI.askMail('${id}')">Queue</button></div>`}
      <div class="mailbody" style="margin-top:10px">${sanitize(m.body.content)}</div>`);
      if (!draft && !m.isRead) OMNI.mailAct(id, 'read');
    } catch (e) { sheet(err(e)); } },
  compose(id, kind) { const c = document.getElementById('composer'); c.innerHTML = `<div class="card"><div class="t">${kind === 'replyAll' ? 'Reply all' : 'Reply'}</div><textarea id="rbody" rows="6" placeholder="Type your reply"></textarea>
      <div class="acts"><button class="btn pri" onclick="OMNI.sendReply('${id}','${kind}',true,this)">Send</button><button class="btn" onclick="OMNI.sendReply('${id}','${kind}',false,this)">Save as draft</button></div></div>`; document.getElementById('rbody').focus(); },
  async sendReply(id, kind, now, btn) { const t = document.getElementById('rbody').value.trim(); if (!t) return toast('Type something first'); btn.disabled = true;
    const comment = t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/\n/g, '<br>');
    try { if (now) await gw(`/me/messages/${id}/${kind}`, {body: {comment}}, MAIL_SCOPES, 'Sending email');
      else await gw(`/me/messages/${id}/${kind === 'replyAll' ? 'createReplyAll' : 'createReply'}`, {body: {comment}}, MAIL_SCOPES, 'Saving drafts');
      toast(now ? 'Sent ✓' : 'Saved to Drafts'); delete mem.inbox; delete mem.drafts; $('#sheet').hidden = true; } catch (e) { btn.disabled = false; toast(e.message); } },
  async sendDraft(id, btn) { if (btn) btn.disabled = true; try { await gw(`/me/messages/${id}/send`, {}, MAIL_SCOPES, 'Sending email'); toast('Sent ✓'); delete mem.drafts; $('#sheet').hidden = true; render(); } catch (e) { if (btn) btn.disabled = false; toast(e.message); } },
  async mailAct(id, act, btn) { try {
      if (act === 'archive') await gw(`/me/messages/${id}/move`, {body: {destinationId: 'archive'}}, MAIL_SCOPES, 'Archiving');
      else await gw(`/me/messages/${id}`, {method: 'PATCH', body: {isRead: act === 'read'}}, MAIL_SCOPES, 'Marking read');
      delete mem.inbox; if (btn) { toast(act === 'archive' ? 'Archived' : 'Done'); if (act === 'archive') { $('#sheet').hidden = true; render(); } } } catch (e) { if (btn) toast(e.message); } },
  async askMail(id) { const t = document.getElementById('askc').value.trim(); if (!t) return toast('Say what you want done');
    try { const m = await g(`/me/messages/${id}?$select=internetMessageId,subject`); await queueAdd({kind: 'email', ref: m.internetMessageId || id, title: 'Email: ' + (m.subject || ''), text: t, project: 'Org Instructions', projectId: '01a123f3-23f0-74f4-ba0d-6982b034e16d'}); toast('Queued for Claude'); document.getElementById('askc').value = ''; } catch (e) { toast(e.message); } },
  async ask() { const t = document.getElementById('qtext').value.trim(); if (!t) return toast('Type a request'); const p = document.getElementById('qproj').value; const pr = PROJECTS.find(x => x[0] === p) || [];
    try { await queueAdd({kind: 'ask', title: t.slice(0, 80), text: t, project: p, projectId: pr[1] || ''}); toast('Queued for Claude'); render(); } catch (e) { toast(e.message); } },
  async answer(qid) { const t = document.getElementById('ans_' + qid).value.trim(); if (!t) return;
    try { await queueWrite(d => { const i = d[qid]; if (i) { i.text = (i.text || '') + '\n\nTom (' + new Date().toLocaleString() + '): ' + t; i.status = 'new'; i._at = new Date().toISOString(); } }); toast('Sent back to Claude'); render(); } catch (e) { toast(e.message); } },
  async decide(fid, status) { const path = C.paths.boards + '/content/findings.json';
    try { const j = safeJSON(await spText(path)); const f = j.docs[fid]; if (!f) throw new Error('Card not found'); const now = new Date().toISOString();
      f.status = status; f.decidedAt = now; f.decidedBy = 'tom (phone app)'; f._at = now; j.updatedAt = now;
      await gw(spItem(path) + '/content', {method: 'PUT', body: JSON.stringify(j)}, WRITE_SCOPES, 'Saving'); delete mem.findings; toast(status === 'approved' ? 'Approved' : 'Skipped'); render(); } catch (e) { toast(e.message); } },
  async report(id) { sheet('<div class="loading">Opening report…</div>');
    try { const a = await g(`/me/messages/${id}/attachments?$select=id,name,contentType,size`); const files = (a.value || []).filter(x => !x.isInline);
      if (!files.length) return sheet('<div class="empty">No attachment on this email.</div>');
      let h = ''; for (const f of files) { h += `<h2>${esc(f.name)}</h2>`;
        if (/\.(xlsx|xls|csv)$/i.test(f.name)) { const buf = DEMO ? await g(`/att/${f.id}`) : await (await g(`/me/messages/${id}/attachments/${f.id}/$value`, {raw: true})).arrayBuffer(); const wb = await readXlsx(buf);
          for (const [sn, rows] of Object.entries(wb)) { const body = rows.filter(r => r.some(c => String(c).trim())); h += `<div class="muted">${esc(sn)} · ${Math.max(0, body.length - 1)} rows</div><div class="tbl"><table>${body.slice(0, 300).map((r, i) => `<tr>${r.map(c => i === 0 ? `<th>${esc(c)}</th>` : `<td>${esc(c)}</td>`).join('')}</tr>`).join('')}</table></div>`; } }
        else { const r = await g(`/me/messages/${id}/attachments/${f.id}/$value`, {raw: true}); const url = URL.createObjectURL(await r.blob()); h += `<a class="btn pri" href="${url}" target="_blank">Open ${esc(f.name)}</a>`; } }
      sheet(h); } catch (e) { sheet(err(e)); } },
  post(id) { const p = OMNI._posts[id] || {}; sheet(`<div class="t" style="font-size:17px;padding-right:30px">${esc(p.title)}</div><div class="s">${esc(p.channel)} · ${esc(p.date)} · ${esc(p.status)}</div>${p.photoUrl ? `<img src="${esc(p.photoUrl)}" style="width:100%;border-radius:10px;margin-top:10px" alt="">` : ''}<div class="p">${esc(p.body || p.text || p.caption || '')}</div>${p.hashtags ? `<div class="s">${esc([].concat(p.hashtags).join(' '))}</div>` : ''}`); }
};
Object.assign(window.OMNI, OMNI_X);
function sanitize(html) {
  const d = new DOMParser().parseFromString(html || '', 'text/html');
  d.querySelectorAll('script,style,iframe,object,embed,link,meta,form').forEach(n => n.remove());
  d.querySelectorAll('*').forEach(n => [...n.attributes].forEach(a => { if (/^on/i.test(a.name) || (/^(href|src)$/i.test(a.name) && /^\s*javascript:/i.test(a.value))) n.removeAttribute(a.name); }));
  d.querySelectorAll('a').forEach(a => { a.target = '_blank'; a.rel = 'noopener'; });
  return d.body.innerHTML;
}

/* ---------------- setup screen + boot ---------------- */
function setupScreen() {
  const redirect = location.origin + location.pathname;
  view.innerHTML = `<div class="setup"><h2>One-time setup</h2><div class="card"><div class="t">Connect to Omni's Microsoft 365</div>
  <div class="s">The app signs in with your own Omni login and reads your Outlook, To Do, calendar and the SharePoint board files. It needs one Microsoft app registration (an IT admin step, once for the whole company).</div>
  <ol class="p"><li>entra.microsoft.com → App registrations → New registration</li><li>Name <b>Omni Command Center</b>, "this organizational directory only"</li><li>Redirect URI: platform <b>Single-page application</b>, URL <b>${esc(redirect)}</b></li>
  <li>API permissions → Microsoft Graph → Delegated: ${SCOPES.map(esc).join(', ')} → Grant admin consent</li><li>Copy the <b>Application (client) ID</b> below</li></ol>
  <div class="addrow"><input id="cid" placeholder="Application (client) ID"><button class="btn pri" id="savecid">Save</button></div></div></div>`;
  $('#savecid').onclick = () => { const v = $('#cid').value.trim(); if (!/^[0-9a-f-]{36}$/i.test(v)) return toast('That isn\'t a client ID'); store.set('clientId', v); location.reload(); };
}
function signInScreen() {
  view.innerHTML = `<div class="setup center"><div class="dot" style="width:18px;height:18px"></div><h2 style="margin-top:14px">Omni Command Center</h2><div class="muted">Sign in with your Omni Microsoft account</div>
  <p><button class="btn pri big" id="signin">Sign in</button></p></div>`;
  $('#signin').onclick = () => pca.loginRedirect({scopes: SCOPES});
}
async function boot() {
  if ('serviceWorker' in navigator && !DEMO && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
  if (DEMO) { await new Promise(r => { const s = document.createElement('script'); s.src = 'mock.js'; s.onload = r; document.head.appendChild(s); }); }
  try {
    if (!(await initAuth())) return setupScreen();
    if (!DEMO && !account) return signInScreen();
    render();
  } catch (e) { view.innerHTML = err(e); }
}
document.addEventListener('visibilitychange', () => { if (!document.hidden && (account || DEMO)) { const last = store.get('lastShow', 0); if (Date.now() - last > 5 * 6e4) { Object.keys(mem).forEach(k => delete mem[k]); render(); } store.set('lastShow', Date.now()); } });
boot();
})();
