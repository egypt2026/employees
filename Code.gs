// الصق هذا الملف كاملاً في Google Apps Script المرتبط بجدول Google Sheets
const ADMIN_KEY = 'غيّر-كلمة-السر-هنا';
const DEF = {start:'09:00',end:'17:00',grace:10,mult:1,weekend:[5],cur:'ج.م',ot:1.5};
const toMin = t => { const p = t.split(':'); return +p[0]*60 + +p[1]; };
const rows = x => { const v = x.getDataRange().getDisplayValues(); v.shift(); return v; };
function sh(n, h) {
  const s = SpreadsheetApp.getActive(); let x = s.getSheetByName(n);
  if (!x) { x = s.insertSheet(n); x.getRange('A:E').setNumberFormat('@'); x.appendRow(h); }
  return x;
}
function now() {
  const d = new Date(), z = 'Africa/Cairo';
  return {d: Utilities.formatDate(d, z, 'yyyy-MM-dd'), t: Utilities.formatDate(d, z, 'HH:mm')};
}
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
function doGet() { return out({ok: 1}); }
function doPost(e) {
  const L = LockService.getScriptLock(); L.waitLock(20000);
  try { return out(handle(JSON.parse(e.postData.contents))); }
  catch (x) { return out({err: String(x.message || x)}); }
  finally { L.releaseLock(); }
}
function handle(q) {
  const n = now(), E = sh('Employees', ['id','name','pin','salary','fp']), G = sh('Logs', ['date','id','in','out']), C = sh('Config', ['json']);
  if (!E.getRange(1, 5).getValue()) E.getRange(1, 5).setValue('fp');
  if (!E.getRange(1, 6).getValue()) E.getRange(1, 6, 1, 2).setValues([['st', 'en']]);
  const cfg = () => { const r = rows(C); return Object.assign({}, DEF, r[0] && r[0][0] ? JSON.parse(r[0][0]) : {}); };
  const cache = CacheService.getScriptCache();
  const guard = (k, max) => { if ((+cache.get(k) || 0) >= max) throw new Error('تم إيقاف المحاولات مؤقتاً لمدة 10 دقائق'); };
  const fail = k => cache.put(k, (+cache.get(k) || 0) + 1, 600);
  const log = r => ({d: r[0], id: r[1], in: r[2], out: r[3]});

  if (q.a == 'pub') return {now: n, cfg: cfg(), emps: rows(E).map(r => ({id: r[0], name: r[1], st: r[5], en: r[6]})), logs: rows(G).filter(r => r[0] == n.d).map(log)};

  if (q.a == 'punch') {
    const k = 'p' + q.id; guard(k, 5);
    const e = rows(E).find(r => r[0] == q.id);
    if (!e || e[2] !== String(q.pin)) { fail(k); throw new Error('الرقم السري غير صحيح'); }
    cache.remove(k);
    const g = rows(G), i = g.findIndex(r => r[0] == n.d && r[1] == q.id);
    if (i < 0) { G.appendRow([n.d, q.id, n.t, '']); return {kind: 'in', t: n.t}; }
    if (!g[i][3]) {
      if (toMin(n.t) - toMin(g[i][2]) < 2) return {kind: 'dup', t: n.t};
      G.getRange(i + 2, 4).setValue(n.t); return {kind: 'out', t: n.t};
    }
    return {kind: 'done', t: n.t};
  }

  guard('a', 10);
  if (q.key !== ADMIN_KEY) { fail('a'); throw new Error('كلمة سر الإدارة غير صحيحة'); }
  cache.remove('a');
  if (q.a == 'adm_get') return {cfg: cfg(), emps: rows(E).map(r => ({id: r[0], name: r[1], pin: r[2], salary: +r[3], fp: r[4], st: r[5], en: r[6]})), logs: rows(G).map(log)};
  if (q.a == 'adm_add') { E.appendRow([Date.now().toString(36), q.name, String(q.pin), String(q.salary), String(q.fp || '')]); return {ok: 1}; }
  if (q.a == 'adm_del') { const i = rows(E).findIndex(r => r[0] == q.id); if (i >= 0) E.deleteRow(i + 2); return {ok: 1}; }
  if (q.a == 'adm_log') {
    const g = rows(G), i = g.findIndex(r => r[0] == q.d && r[1] == q.id);
    if (!q.in) { if (i >= 0) G.deleteRow(i + 2); }
    else if (i >= 0) G.getRange(i + 2, 3, 1, 2).setValues([[q.in, q.out]]);
    else G.appendRow([q.d, q.id, q.in, q.out]);
    return {ok: 1};
  }
  if (q.a == 'adm_import') {
    const g = rows(G), idx = {}, add = {};
    g.forEach((r, i) => idx[r[0] + '|' + r[1]] = i);
    q.rows.forEach(x => {
      const k = x.d + '|' + x.id;
      if (k in idx) G.getRange(idx[k] + 2, 3, 1, 2).setValues([[x.in, x.out || '']]);
      else add[k] = [x.d, x.id, x.in, x.out || ''];
    });
    const a = Object.keys(add).map(k => add[k]);
    if (a.length) G.getRange(G.getLastRow() + 1, 1, a.length, 4).setValues(a);
    return {n: q.rows.length};
  }
  if (q.a == 'adm_fp') { const i = rows(E).findIndex(r => r[0] == q.id); if (i >= 0) E.getRange(i + 2, 5).setValue(String(q.fp)); return {ok: 1}; }
  if (q.a == 'adm_emp') {
    const i = rows(E).findIndex(r => r[0] == q.id); if (i < 0) throw new Error('الموظف غير موجود');
    const R = E.getRange(i + 2, 1, 1, 7); R.setNumberFormat('@');
    R.setValues([[q.id, q.name, String(q.pin), String(q.salary), String(q.fp || ''), q.st || '', q.en || '']]);
    return {ok: 1};
  }
  if (q.a == 'adm_cfg') { C.getRange(2, 1).setValue(JSON.stringify(q.cfg)); return {ok: 1}; }
  throw new Error('طلب غير معروف');
}
