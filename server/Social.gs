// תוספת: לייקים, תגובות וקטגוריות — העוגות של החיים
HEAD.Likes = ['code', 'id', 'memberId', 'name', 'ts'];
HEAD.Comments = ['code', 'id', 'cid', 'memberId', 'name', 'text', 'ts', 'deleted'];
if (HEAD.Families.indexOf('cats') < 0) HEAD.Families.push('cats');

function doPost(e) {
  let res;
  try {
    const q = JSON.parse(e.postData.contents);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try { res = route2_(q); } finally { lock.releaseLock(); }
  } catch (err) {
    res = { ok: false, error: String((err && err.message) || err) };
  }
  return out_(res);
}

function head2_(n) {
  const s = sheet_(n);
  if (s.getLastColumn() < HEAD[n].length) s.getRange(1, 1, 1, HEAD[n].length).setValues([HEAD[n]]);
  return s;
}

function route2_(q) {
  const a = q.action;
  const extra = ['list', 'like', 'comment', 'delComment', 'setCats'];
  if (extra.indexOf(a) < 0) return route_(q);
  head2_('Families');
  if (a === 'list') {
    const res = route_(q);
    const code = String(q.code || '').trim().toUpperCase();
    const me = rows_('Members').find(m => String(m.code) === code && String(m.memberId) === String(q.memberId));
    const fam = rows_('Families').find(f => String(f.code) === code);
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
    res.social = social; res.cats = cats;
    return res;
  }
  const code = String(q.code || '').trim().toUpperCase();
  const fam = rows_('Families').find(f => String(f.code) === code);
  if (!fam) fail_('קוד משפחה לא נמצא');
  const me = rows_('Members').find(m => String(m.code) === code && String(m.memberId) === String(q.memberId));
  if (!me) fail_('לא מחובר למשפחה — הצטרפו מחדש');
  const admin = me.role === 'admin';

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
}
