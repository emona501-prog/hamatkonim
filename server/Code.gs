/**
 * העוגות של החיים — שרת משפחתי
 * Google Apps Script + Google Sheets
 * חי אפליקציות
 *
 * התקנה: פותחים script.google.com/create, מדביקים את כל הקוד הזה,
 * (הגיליון עם הנתונים נוצר לבד בדרייב), ופורסים כ-Web app (Execute as: Me, Who has access: Anyone).
 */

const APP = 'העוגות של החיים';
const HEAD = {
  Families: ['code', 'name', 'pin', 'created', 'cats'],
  Members: ['code', 'memberId', 'name', 'role', 'joined'],
  Recipes: ['code', 'id', 'status', 'by', 'byName', 'updated', 'deleted', 'data'],
  Likes: ['code', 'id', 'memberId', 'name', 'ts'],
  Comments: ['code', 'id', 'cid', 'memberId', 'name', 'text', 'ts', 'deleted']
};

function doGet() {
  return out_({ ok: true, app: APP });
}

function doPost(e) {
  let res;
  try {
    const q = JSON.parse(e.postData.contents);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try { res = route_(q); } finally { lock.releaseLock(); }
  } catch (err) {
    res = { ok: false, error: String((err && err.message) || err) };
  }
  return out_(res);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- sheets ---------- */
function ss_() {
  let a = null;
  try { a = SpreadsheetApp.getActive(); } catch (e) {}
  if (a) return a;
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('SS');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const s = SpreadsheetApp.create(APP + ' - נתונים');
  p.setProperty('SS', s.getId());
  return s;
}
function sheet_(n) {
  const ss = ss_();
  let s = ss.getSheetByName(n);
  if (!s) { s = ss.insertSheet(n); s.appendRow(HEAD[n]); s.setFrozenRows(1); return s; }
  const w = s.getLastColumn();
  if (w < HEAD[n].length) s.getRange(1, 1, 1, HEAD[n].length).setValues([HEAD[n]]);
  return s;
}
function rows_(n) {
  const v = sheet_(n).getDataRange().getValues();
  const h = v.shift() || [];
  return v.map((r, i) => { const o = { _row: i + 2 }; h.forEach((k, j) => o[k] = r[j]); return o; });
}
function add_(n, o) {
  sheet_(n).appendRow(HEAD[n].map(k => (o[k] === undefined ? '' : o[k])));
}
function set_(n, row, o) {
  const s = sheet_(n);
  HEAD[n].forEach((k, j) => { if (k in o) s.getRange(row, j + 1).setValue(o[k]); });
}
function txt_(v) {
  const s = String(v == null ? '' : v).trim().slice(0, 80);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function rand_(n) {
  const c = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < n; i++) s += c[Math.floor(Math.random() * c.length)];
  return s;
}
function fail_(m) { throw new Error(m); }

/* ---------- images → Google Drive ---------- */
function folder_(code) {
  return sub_(sub_(DriveApp.getRootFolder(), APP + ' - תמונות'), code);
}
function sub_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function img_(code, d) {
  if (typeof d !== 'string') return d;
  const m = /^data:(image\/[\w+.-]+);base64,(.*)$/.exec(d);
  if (!m) return d;
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'img_' + Date.now() + '.jpg');
  const f = folder_(code).createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1200';
}

/* ---------- api ---------- */
function route_(q) {
  const a = q.action;
  const code = String(q.code || '').trim().toUpperCase();

  if (a === 'ping') return { ok: true, app: APP };

  if (a === 'create') {
    const fname = txt_(q.familyName), name = txt_(q.name), pin = String(q.pin || '').trim();
    if (!fname || !name) fail_('חסר שם המשפחה או שם המנהלת');
    if (!/^\d{4,8}$/.test(pin)) fail_('סיסמת מנהלת: 4 עד 8 ספרות');
    const fams = rows_('Families');
    let c;
    do { c = rand_(6); } while (fams.some(f => String(f.code) === c));
    add_('Families', { code: c, name: fname, pin: "'" + pin, created: Date.now() });
    const memberId = rand_(16);
    add_('Members', { code: c, memberId, name, role: 'admin', joined: Date.now() });
    return { ok: true, code: c, memberId, role: 'admin', familyName: fname };
  }

  const fam = rows_('Families').find(f => String(f.code) === code);
  if (!fam) fail_('קוד משפחה לא נמצא');

  if (a === 'join') {
    const name = txt_(q.name);
    if (!name) fail_('חסר שם');
    const memberId = rand_(16);
    add_('Members', { code, memberId, name, role: 'member', joined: Date.now() });
    return { ok: true, code, memberId, role: 'member', familyName: fam.name };
  }

  if (a === 'adminLogin') {
    if (String(fam.pin) !== String(q.pin || '').trim()) fail_('סיסמה שגויה');
    const name = txt_(q.name) || 'מנהלת';
    const memberId = rand_(16);
    add_('Members', { code, memberId, name, role: 'admin', joined: Date.now() });
    return { ok: true, code, memberId, role: 'admin', familyName: fam.name };
  }

  const me = rows_('Members').find(m => String(m.code) === code && String(m.memberId) === String(q.memberId));
  if (!me) fail_('לא מחובר למשפחה — הצטרפו מחדש');
  const admin = me.role === 'admin';
  const all = () => rows_('Recipes').filter(r => String(r.code) === code);

  if (a === 'list') {
    const since = Number(q.since) || 0;
    const recs = all();
    const list = recs.filter(r => Number(r.updated) > since)
      .filter(r => r.deleted === true || String(r.deleted) === 'true' || r.status === 'approved' || admin || String(r.by) === me.memberId)
      .map(r => {
        const del = r.deleted === true || String(r.deleted) === 'true';
        let data = null;
        if (!del) { try { data = JSON.parse(r.data); } catch (e) { data = null; } }
        return { id: String(r.id), status: r.status, by: String(r.by), byName: String(r.byName), updated: Number(r.updated), deleted: del, data };
      });
    const pending = admin ? recs.filter(r => r.status === 'pending' && !(r.deleted === true || String(r.deleted) === 'true')).length : 0;
    const social = {};
    const S = id => social[id] || (social[id] = { likes: 0, mine: false, likers: [], comments: [] });
    rows_('Likes').filter(l => String(l.code) === code).forEach(l => {
      const o = S(String(l.id)); o.likes++; o.likers.push(String(l.name));
      if (String(l.memberId) === me.memberId) o.mine = true;
    });
    rows_('Comments').filter(c => String(c.code) === code && !(c.deleted === true || String(c.deleted) === 'true')).forEach(c => {
      S(String(c.id)).comments.push({ cid: String(c.cid), name: String(c.name), text: String(c.text), ts: Number(c.ts), mine: String(c.memberId) === me.memberId });
    });
    let cats = [];
    try { cats = JSON.parse(fam.cats || '[]'); } catch (e) {}
    return { ok: true, recipes: list, now: Date.now(), pending, familyName: fam.name, role: me.role, name: me.name, social, cats };
  }

  if (a === 'submit' || a === 'submitMany') {
    const arr = a === 'submit' ? [q.recipe] : (q.recipes || []);
    const recs = all();
    const results = arr.map(r => {
      if (!r || !r.id) fail_('מתכון לא תקין');
      ['have', 'done', 'fam', 'sample', 'seed'].forEach(k => delete r[k]);
      if (r.dish) r.dish = img_(code, r.dish);
      if (r.pic) r.pic = img_(code, r.pic);
      if (Array.isArray(r.scans)) r.scans = r.scans.map(s => img_(code, s));
      const data = JSON.stringify(r);
      if (data.length > 49000) fail_('המתכון "' + (r.title || '') + '" ארוך מדי');
      const ex = recs.find(x => String(x.id) === String(r.id));
      const now = Date.now();
      if (ex) {
        const mine = String(ex.by) === me.memberId;
        if (!admin && !(mine && ex.status !== 'approved')) fail_('רק המנהלת יכולה לערוך מתכון מאושר');
        const status = admin ? 'approved' : 'pending';
        set_('Recipes', ex._row, { status, updated: now, deleted: false, data });
        return { id: r.id, status, updated: now, recipe: r };
      }
      const status = admin ? 'approved' : 'pending';
      add_('Recipes', { code, id: String(r.id), status, by: me.memberId, byName: me.name, updated: now, deleted: false, data });
      return { id: r.id, status, updated: now, recipe: r };
    });
    return { ok: true, results };
  }

  if (a === 'approve' || a === 'reject' || a === 'delete') {
    const ex = all().find(x => String(x.id) === String(q.id));
    if (!ex) fail_('המתכון לא נמצא');
    const mine = String(ex.by) === me.memberId;
    if (a === 'delete') {
      if (!admin && !(mine && ex.status !== 'approved')) fail_('רק המנהלת יכולה למחוק');
      set_('Recipes', ex._row, { deleted: true, updated: Date.now() });
    } else {
      if (!admin) fail_('רק המנהלת יכולה לאשר');
      set_('Recipes', ex._row, { status: a === 'approve' ? 'approved' : 'rejected', updated: Date.now() });
    }
    return { ok: true };
  }

  if (a === 'like') {
    const id = String(q.id);
    const ex = rows_('Likes').find(l => String(l.code) === code && String(l.id) === id && String(l.memberId) === me.memberId);
    if (ex) sheet_('Likes').deleteRow(ex._row);
    else add_('Likes', { code, id, memberId: me.memberId, name: me.name, ts: Date.now() });
    return { ok: true, liked: !ex };
  }

  if (a === 'comment') {
    const text = String(q.text || '').trim().slice(0, 1000);
    if (!text) fail_('התגובה ריקה');
    const t = /^[=+\-@]/.test(text) ? "'" + text : text;
    const cid = rand_(10);
    add_('Comments', { code, id: String(q.id), cid, memberId: me.memberId, name: me.name, text: t, ts: Date.now(), deleted: false });
    return { ok: true, cid };
  }

  if (a === 'delComment') {
    const c = rows_('Comments').find(x => String(x.code) === code && String(x.cid) === String(q.cid));
    if (!c) fail_('התגובה לא נמצאה');
    if (!admin && String(c.memberId) !== me.memberId) fail_('אפשר למחוק רק תגובה שלך');
    set_('Comments', c._row, { deleted: true });
    return { ok: true };
  }

  if (a === 'setCats') {
    if (!admin) fail_('רק המנהלים יכולים לשנות קטגוריות');
    const cats = (q.cats || []).map(c => String(c).trim().slice(0, 30)).filter(Boolean).slice(0, 40);
    set_('Families', fam._row, { cats: JSON.stringify(cats) });
    return { ok: true, cats };
  }

  if (a === 'members') {
    if (!admin) fail_('רק המנהלת');
    return { ok: true, members: rows_('Members').filter(m => String(m.code) === code).map(m => ({ name: String(m.name), role: m.role, joined: Number(m.joined) })) };
  }

  fail_('פעולה לא מוכרת');
}
