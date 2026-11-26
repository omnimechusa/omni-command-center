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
  qbo: () => data('qbolive', () => spJSON(board('finance/qbo/live.json')), 30),
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

/* ---------------- router ---------------- */
const TABS = {
  command: [['today', 'Today'], ['todo', 'To Do'], ['admin', 'Admin To Do'], ['inbox', 'Inbox'], ['queue', 'Claude queue']],
  construction: [['checkins', 'Daily Check-Ins'], ['wrap', 'Wrap-Up'], ['jobs', 'Job Board'], ['thursday', 'Thursday Reports'], ['permits', 'Permits'], ['permitdesk', 'Permit Desk']],
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
  const my = ++seq;
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
  let h = `<h2>Good ${new Date().getHours() < 12 ? 'morning' : 'afternoon'}${me ? ', ' + esc((me.displayName || '').split(' ')[0]) : ''}<small>${esc(day(new Date()))}</small></h2>
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
V.command_admin = async () => {
  const cfg = store.get('admin', null);
  const lists = (await L.lists().catch(() => ({value: []}))).value || [];
  if (cfg && cfg.type === 'todo') { const l = lists.find(x => x.id === cfg.id); if (l) return todoView(l, lists, null, true); }
  if (cfg && cfg.type === 'link') return `<h2>Admin To Do</h2><div class="card"><div class="t">${esc(cfg.name || 'Admin To Do')}</div><div class="s">${esc(cfg.url)}</div><div class="acts"><a class="btn pri" href="${esc(cfg.url)}" target="_blank" rel="noopener">Open</a><button class="btn" onclick="OMNI.adminReset()">Change</button></div></div>`;
  return `<h2>Admin To Do<small>set once</small></h2><div class="card"><div class="t">Which list is it?</div><div class="s">Pick the To Do list, or paste the link of the app you saved from Chrome.</div>
  <div class="p">${lists.map(l => `<button class="btn" style="margin:4px 4px 0 0" onclick="OMNI.adminSet('todo','${l.id}')">${esc(l.displayName)}</button>`).join('')}</div>
  <div class="addrow" style="margin-top:12px"><input id="adminUrl" placeholder="https://… (app link)"><button class="btn pri" onclick="OMNI.adminSet('link',document.getElementById('adminUrl').value)">Save</button></div></div>`;
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
V.command_queue = async () => {
  const q = await L.queue(); const items = Array.isArray(q) ? q : Array.isArray(q.items) ? q.items : q.docs ? Object.entries(q.docs).map(([id, v]) => Object.assign({id}, v)) : [];
  items.sort((a, b) => String(b.created || b.createdAt || b.at || '').localeCompare(String(a.created || a.createdAt || a.at || '')));
  const st = s => /done|ok|complete/i.test(s) ? 'g' : /error|fail/i.test(s) ? 'r' : /need|wait|working|new/i.test(s) ? 'y' : '';
  return `<h2>Claude queue<small>${items.length} items</small></h2>` + (items.slice(0, 60).map(i => `<div class="card"><div class="row"><div class="t">${esc(i.title || i.subject || i.text || i.prompt || i.id)}</div><span class="pill ${st(i.status || '')}">${esc(i.status || '')}</span></div>
    <div class="s">${esc(when(i.created || i.createdAt || i.at))}${i.project ? ' · ' + esc(i.project) : ''}</div>${i.result || i.note ? `<div class="p muted">${esc(String(i.result || i.note).slice(0, 300))}</div>` : ''}</div>`).join('') || '<div class="empty">Queue is empty.</div>')
    + `<div class="acts"><a class="btn" href="${C.links.claudeBoard}" target="_blank" rel="noopener">Open full board in Claude</a></div>`;
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
  let h = `<div class="muted">QuickBooks as of ${esc(dfmt(q.asOf, {month: 'numeric', day: 'numeric', hour: 'numeric', minute: '2-digit'}))}</div>
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
V.marketing_uniforms = async () => {
  let orders = null; try { orders = await L.uniforms(); } catch (_) {}
  const reps = await L.uniformReports().catch(() => ({}));
  const lines = orders ? (Array.isArray(orders) ? orders : orders.lines || Object.values(orders.docs || {})) : [];
  const open = lines.filter(l => !/ordered|delivered|done|cancel/i.test(l.status || ''));
  let h = `<div class="kpis">${kpi(orders ? lines.length : '–', 'Order lines')}${kpi(orders ? open.length : '–', 'Not ordered yet', open.length ? 'hot' : '')}${kpi(orders ? money(lines.reduce((s, l) => s + num(l.price) * (num(l.qty) || 1), 0)) : '–', 'Order value')}${kpi(orders ? new Set(lines.map(l => l.employee || l.name)).size : '–', 'People')}</div>
  <div class="acts"><a class="btn pri" href="${C.links.uniformApp}" target="_blank" rel="noopener">Open uniform order app</a></div>`;
  if (!orders) h += `<div class="card"><div class="s">Orders live inside the uniform app. They show here once the app saves a copy to SharePoint (boards/marketing/uniform/orders.json).</div></div>`;
  else h += '<h2>Open orders</h2>' + (open.map(l => `<div class="card"><div class="row"><div class="t">${esc(l.employee || l.name)}</div><div class="amt">${esc(l.qty || 1)} × ${esc(l.item || l.garment || l.key)}</div></div><div class="s">${esc(l.size || '')} ${esc(l.dept || l.department || '')} · ${esc(when(l.submittedAt))}</div></div>`).join('') || '<div class="empty">Nothing open.</div>');
  const rs = (reps.value || reps || []).filter ? (reps.value || reps) : [];
  h += '<h2>Uniform reports</h2>' + ((Array.isArray(rs) ? rs : []).slice(0, 8).map(m => `<div class="card"><div class="row"><div class="t">${esc(m.subject)}</div><div class="s">${esc(day(m.receivedDateTime))}</div></div><div class="acts">${m.hasAttachments ? `<button class="btn pri" onclick="OMNI.report('${m.id}')">Open report</button>` : ''}<a class="btn" href="${esc(m.webLink)}" target="_blank" rel="noopener">Email</a></div></div>`).join('') || '<div class="empty">No uniform report emails found.</div>');
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

/* ---------------- actions ---------------- */
window.OMNI = {
  go, tool: openTool, set(k, v) { store.set(k, v); render(); },
  pickList(k, id) { store.set(k, id); render(); },
  adminSet(type, v) { if (type === 'link' && !/^https?:\/\//.test(v)) return toast('Paste the full https:// link'); store.set('admin', type === 'todo' ? {type, id: v} : {type, url: v, name: 'Admin To Do'}); render(); },
  adminReset() { store.set('admin', null); render(); },
  async done(btn, list, id) { btn.classList.add('done'); try { await g(`/me/todo/lists/${list}/tasks/${id}`, {method: 'PATCH', body: {status: 'completed'}}); delete mem['tasks.' + list]; btn.closest('.card').style.opacity = .35; toast('Done ✓'); } catch (e) { btn.classList.remove('done'); toast(e.message); } },
  async addTask(list) { const i = document.getElementById('newTask'); const t = i.value.trim(); if (!t) return; i.value = '';
    try { await g(`/me/todo/lists/${list}/tasks`, {method: 'POST', body: {title: t, dueDateTime: {dateTime: todayYMD() + 'T17:00:00', timeZone: 'Pacific Standard Time'}}}); delete mem['tasks.' + list]; render(); toast('Added'); } catch (e) { toast(e.message); } },
  async mail(id) { sheet('<div class="loading">Opening…</div>');
    try { const m = await g(`/me/messages/${id}?$select=subject,from,receivedDateTime,body,webLink,toRecipients`);
      sheet(`<div class="t" style="font-size:17px;padding-right:30px">${esc(m.subject)}</div><div class="s">${esc(m.from && m.from.emailAddress.name)} · ${esc(dfmt(m.receivedDateTime, {month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit'}))}</div>
      <div class="acts"><a class="btn pri" href="${esc(m.webLink)}" target="_blank" rel="noopener">Open in Outlook</a></div><div class="mailbody" style="margin-top:10px">${sanitize(m.body.content)}</div>`); } catch (e) { sheet(err(e)); } },
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
