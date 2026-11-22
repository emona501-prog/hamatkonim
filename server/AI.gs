// תוספת: פענוח מתכונים עם Gemini דרך השרת — העוגות של החיים
// המפתח נשמר בהגדרות הסקריפט (Script Properties) ולעולם לא נשלח לטלפונים.
const SUBMIT_CODE_ = '1988';

function doPost(e) {
  let res;
  try {
    const q = JSON.parse(e.postData.contents);
    if (q.action === 'parse') {
      res = route3_(q);
    } else {
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try { res = route3_(q); } finally { lock.releaseLock(); }
    }
  } catch (err) {
    res = { ok: false, error: String((err && err.message) || err) };
  }
  return out_(res);
}

function key_() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_KEY') || '';
}

function route3_(q) {
  const a = q.action;
  if (['parse', 'setKey', 'list'].indexOf(a) < 0) return route2_(q);
  if (a === 'list') {
    const r = route2_(q);
    if (r && r.ok) r.hasAi = !!key_();
    return r;
  }
  const code = String(q.code || '').trim().toUpperCase();
  const me = rows_('Members').find(m => String(m.code) === code && String(m.memberId) === String(q.memberId));
  if (!me) fail_('לא מחובר למשפחה — הצטרפו מחדש');
  const admin = me.role === 'admin';

  if (a === 'setKey') {
    if (!admin) fail_('רק מנהלים יכולים לשמור מפתח');
    const k = String(q.key || '').trim();
    if (k && !/^AIza[\w-]{20,}$/.test(k)) fail_('המפתח לא נראה תקין (צריך להתחיל ב-AIza)');
    if (k) PropertiesService.getScriptProperties().setProperty('GEMINI_KEY', k);
    else PropertiesService.getScriptProperties().deleteProperty('GEMINI_KEY');
    return { ok: true, hasAi: !!k };
  }

  if (a === 'parse') {
    if (!admin && String(q.sc || '') !== SUBMIT_CODE_) fail_('כדי לפענח צריך את קוד שליחת המתכונים');
    const k = key_();
    if (!k) fail_('עוד לא הוגדר מפתח פענוח בשרת');
    const parts = [{ text: String(q.prompt || '').slice(0, 40000) }];
    (q.images || []).slice(0, 8).forEach(d => parts.push({ inline_data: { mime_type: 'image/jpeg', data: String(d) } }));
    const model = String(q.model || 'gemini-2.5-flash').replace(/[^\w.-]/g, '');
    const resp = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + encodeURIComponent(k), {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ contents: [{ parts }], generationConfig: { responseMimeType: 'application/json', temperature: 0.2 } })
    });
    if (resp.getResponseCode() !== 200) fail_('שגיאה מ-Gemini (' + resp.getResponseCode() + ')');
    const j = JSON.parse(resp.getContentText());
    const text = ((j.candidates && j.candidates[0] && j.candidates[0].content && j.candidates[0].content.parts) || []).map(p => p.text || '').join('');
    return { ok: true, text };
  }
}
