// Google Sheet > Extensions > Apps Script এ এই কোড পেস্ট করুন।
// Project Settings > Script Properties এ ADMIN_PASS নামে অ্যাডমিন পাসওয়ার্ড দিন।
const H = {
  Applications: ['আইডি','তারিখ','সেবা','নাম','ফোন','বিবরণ','স্ট্যাটাস','মোট বিল','জমা','নোট'],
  Accounts: ['আইডি','তারিখ','ধরন','বিবরণ','পরিমাণ','আবেদন আইডি']
};
const K = {
  Applications: ['id','date','service','name','phone','details','status','bill','paid','note'],
  Accounts: ['id','date','type','desc','amount','appId']
};
const SERVICES = ['হজ্জ প্রসেসিং','ওমরাহ প্রসেসিং','পাসপোর্ট (নতুন/রিনিউ)','ভিসা প্রসেসিং','টিকেট প্রসেসিং','ল্যান্ড ও হাউজ প্রপার্টিজ'];

function sh(n) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(n);
  if (!s) { s = ss.insertSheet(n); s.appendRow(H[n]); s.setFrozenRows(1); s.getRange(1,1,1,H[n].length).setFontWeight('bold'); }
  return s;
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function today() { return Utilities.formatDate(new Date(), 'Asia/Dhaka', 'yyyy-MM-dd'); }
function rows(n) {
  const v = sh(n).getDataRange().getValues(); v.shift();
  return v.map(r => { const o = {}; K[n].forEach((k, j) => o[k] = r[j] instanceof Date ? Utilities.formatDate(r[j], 'Asia/Dhaka', 'yyyy-MM-dd') : r[j]); return o; });
}
function findRow(n, id) {
  const ids = sh(n).getRange(1, 1, sh(n).getLastRow(), 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(id)); return i < 1 ? 0 : i + 1;
}
function add(n, obj) {
  const s = sh(n), r = s.getLastRow() + 1;
  s.getRange(r, 1, 1, 1).setNumberFormat('@');
  if (n === 'Applications') s.getRange(r, 5).setNumberFormat('@');
  s.getRange(r, 1, 1, K[n].length).setValues([K[n].map(k => obj[k] ?? '')]);
}
function doGet() { return out({ ok: true, msg: 'API চালু আছে' }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    return out(route(JSON.parse(e.postData.contents)));
  } catch (err) { return out({ ok: false, error: String(err.message || err) }); }
  finally { try { lock.releaseLock(); } catch (x) {} }
}

function route(d) {
  const cache = CacheService.getScriptCache();
  // ---- সবার জন্য ----
  if (d.action === 'apply') {
    const name = String(d.name || '').trim().slice(0, 100), phone = String(d.phone || '').trim().slice(0, 20);
    if (!name || !/^[0-9+\- ]{10,15}$/.test(phone)) return { ok: false, error: 'নাম ও সঠিক মোবাইল নম্বর দিন' };
    if (SERVICES.indexOf(d.service) < 0) return { ok: false, error: 'সেবা বেছে নিন' };
    const id = 'A' + Utilities.formatDate(new Date(), 'Asia/Dhaka', 'yyMMdd') + '-' + (sh('Applications').getLastRow());
    add('Applications', { id, date: today(), service: d.service, name, phone, details: String(d.details || '').slice(0, 800), status: 'নতুন', bill: 0, paid: 0, note: '' });
    return { ok: true, id };
  }
  if (d.action === 'track') {
    const a = rows('Applications').find(r => String(r.id).toUpperCase() === String(d.id || '').trim().toUpperCase() && String(r.phone).slice(-6) === String(d.phone || '').trim().slice(-6));
    return a ? { ok: true, service: a.service, status: a.status, date: a.date } : { ok: false, error: 'আইডি ও মোবাইল নম্বর মিলছে না' };
  }
  if (d.action === 'login') {
    const pass = PropertiesService.getScriptProperties().getProperty('ADMIN_PASS');
    if (!pass || d.password !== pass) { Utilities.sleep(1500); return { ok: false, error: 'পাসওয়ার্ড ভুল' }; }
    const t = Utilities.getUuid(); cache.put('t_' + t, '1', 21600); return { ok: true, token: t };
  }
  // ---- শুধু অ্যাডমিন ----
  if (!d.token || !cache.get('t_' + d.token)) return { ok: false, auth: false, error: 'লগইন করুন' };
  if (d.action === 'data') return { ok: true, apps: rows('Applications').reverse(), acc: rows('Accounts').reverse() };
  if (d.action === 'updateApp') {
    const r = findRow('Applications', d.id); if (!r) return { ok: false, error: 'আবেদন পাওয়া যায়নি' };
    const s = sh('Applications');
    s.getRange(r, 7, 1, 4).setValues([[d.status, Number(d.bill) || 0, Number(d.paid) || 0, String(d.note || '').slice(0, 300)]]);
    return { ok: true };
  }
  if (d.action === 'addEntry') {
    if (!d.desc || !(Number(d.amount) > 0)) return { ok: false, error: 'বিবরণ ও পরিমাণ দিন' };
    add('Accounts', { id: 'E' + Date.now(), date: d.date || today(), type: d.type === 'ব্যয়' ? 'ব্যয়' : 'আয়', desc: String(d.desc).slice(0, 200), amount: Number(d.amount), appId: d.appId || '' });
    return { ok: true };
  }
  if (d.action === 'delEntry') {
    const r = findRow('Accounts', d.id); if (r) sh('Accounts').deleteRow(r); return { ok: true };
  }
  return { ok: false, error: 'অজানা অনুরোধ' };
}
