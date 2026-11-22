// Google Sheet > Extensions > Apps Script এ এই কোড পেস্ট করুন (আগের কোড মুছে)।
// Script Properties এ ADMIN_PASS দিন: এটি "মূল অ্যাডমিন" (আইডি: owner) এর পাসওয়ার্ড।
// কোড বদলানোর পর Deploy > Manage deployments > Edit > New version দিয়ে আবার Deploy করুন।
const SCH = {
  Applications: { h: ['আইডি','তারিখ','সেবা','নাম','ফোন','বিবরণ','স্ট্যাটাস','মোট বিল','জমা','নোট'], k: ['id','date','service','name','phone','details','status','bill','paid','note'] },
  Accounts: { h: ['আইডি','তারিখ','ধরন','বিবরণ','পরিমাণ','আবেদন আইডি'], k: ['id','date','type','desc','amount','appId'] },
  Cash: { h: ['আইডি','তারিখ','ধরন','স্থান','পরিমাণ','নোট'], k: ['id','date','type','place','amount','note'] },
  Staff: { h: ['আইডি','নাম','পদবি','মোবাইল','ঠিকানা','যোগদান','মূল বেতন'], k: ['id','name','title','phone','address','joined','salary'] },
  Salaries: { h: ['আইডি','তারিখ','স্টাফ আইডি','স্টাফ','মাস','পরিমাণ','নোট'], k: ['id','date','staffId','staffName','month','amount','note'] },
  Admins: { h: ['আইডি','নাম','পদবি','মোবাইল','ইমেইল','ঠিকানা','হ্যাশ','সল্ট'], k: ['id','name','title','phone','email','address','hash','salt'] },
  Settings: { h: ['কী','মান'], k: ['key','value'] }
};
const SERVICES = ['হজ্জ প্রসেসিং','ওমরাহ প্রসেসিং','পাসপোর্ট (নতুন/রিনিউ)','ভিসা প্রসেসিং','টিকেট প্রসেসিং','ল্যান্ড ও হাউজ প্রপার্টিজ'];
const SET_KEYS = ['siteName','logo','address','phone','email','facebook','instagram','youtube','hours'];
const DEL = { delEntry: 'Accounts', delCash: 'Cash', delStaff: 'Staff', delSalary: 'Salaries' };

function sh(n) {
  const ss = SpreadsheetApp.getActive(); let s = ss.getSheetByName(n);
  if (!s) { s = ss.insertSheet(n); s.appendRow(SCH[n].h); s.setFrozenRows(1); s.getRange(1, 1, 1, SCH[n].h.length).setFontWeight('bold'); }
  return s;
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function today() { return Utilities.formatDate(new Date(), 'Asia/Dhaka', 'yyyy-MM-dd'); }
function rows(n) {
  const v = sh(n).getDataRange().getValues(); v.shift();
  return v.map(r => { const o = {}; SCH[n].k.forEach((k, j) => o[k] = r[j] instanceof Date ? Utilities.formatDate(r[j], 'Asia/Dhaka', 'yyyy-MM-dd') : r[j]); return o; });
}
function findRow(n, id) {
  const s = sh(n), ids = s.getRange(1, 1, s.getLastRow(), 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(id)); return i < 1 ? 0 : i + 1;
}
function add(n, o) {
  const s = sh(n), r = s.getLastRow() + 1, k = SCH[n].k;
  s.getRange(r, 1, 1, k.length).setNumberFormat('@').setValues([k.map(x => o[x] ?? '')]);
}
const tx = (v, l) => String(v || '').trim().slice(0, l);
const amt = v => { v = Number(v); if (!(v > 0)) throw new Error('সঠিক পরিমাণ দিন'); return v; };
const hash = (p, salt) => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + p));
const newId = p => p + Date.now() + Math.floor(Math.random() * 90 + 10);
function settings() { const o = {}; rows('Settings').forEach(r => o[r.key] = r.value); return o; }

function doGet() { return out({ ok: true, msg: 'API চালু আছে' }); }
function doPost(e) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); return out(route(JSON.parse(e.postData.contents))); }
  catch (err) { return out({ ok: false, error: String(err.message || err) }); }
  finally { try { lock.releaseLock(); } catch (x) {} }
}

function route(d) {
  const cache = CacheService.getScriptCache();
  // ---------- সবার জন্য ----------
  if (d.action === 'settings') { const s = settings(), o = {}; SET_KEYS.forEach(k => o[k] = s[k] || ''); return { ok: true, settings: o }; }
  if (d.action === 'apply') {
    const name = tx(d.name, 100), phone = tx(d.phone, 20);
    if (!name || !/^[0-9+\- ]{10,15}$/.test(phone)) return { ok: false, error: 'নাম ও সঠিক মোবাইল নম্বর দিন' };
    if (SERVICES.indexOf(d.service) < 0) return { ok: false, error: 'সেবা বেছে নিন' };
    const id = 'A' + Utilities.formatDate(new Date(), 'Asia/Dhaka', 'yyMMdd') + '-' + sh('Applications').getLastRow();
    add('Applications', { id, date: today(), service: d.service, name, phone, details: tx(d.details, 800), status: 'নতুন', bill: 0, paid: 0, note: '' });
    return { ok: true, id };
  }
  if (d.action === 'track') {
    const a = rows('Applications').find(r => String(r.id).toUpperCase() === String(d.id || '').trim().toUpperCase() && String(r.phone).slice(-6) === String(d.phone || '').trim().slice(-6));
    return a ? { ok: true, service: a.service, status: a.status, date: a.date } : { ok: false, error: 'আইডি ও মোবাইল নম্বর মিলছে না' };
  }
  if (d.action === 'login') {
    const id = tx(d.id, 100).toLowerCase(), pw = String(d.password || ''), owner = PropertiesService.getScriptProperties().getProperty('ADMIN_PASS');
    let who = null;
    if (owner && id === 'owner' && pw === owner) who = { id: 'owner', name: 'মূল অ্যাডমিন' };
    else { const a = rows('Admins').find(x => (String(x.phone).toLowerCase() === id || String(x.email).toLowerCase() === id) && x.hash === hash(pw, x.salt)); if (a) who = { id: a.id, name: a.name }; }
    if (!who) { Utilities.sleep(1500); return { ok: false, error: 'আইডি বা পাসওয়ার্ড ভুল' }; }
    const t = Utilities.getUuid(); cache.put('t_' + t, JSON.stringify(who), 21600);
    return { ok: true, token: t };
  }
  // ---------- শুধু অ্যাডমিন ----------
  const me = d.token && cache.get('t_' + d.token);
  if (!me) return { ok: false, auth: false, error: 'লগইন করুন' };
  const who = JSON.parse(me);

  if (d.action === 'data') {
    return { ok: true, me: who, apps: rows('Applications').reverse(), acc: rows('Accounts').reverse(), cash: rows('Cash').reverse(), staff: rows('Staff'), salaries: rows('Salaries').reverse(),
      admins: rows('Admins').map(a => ({ id: a.id, name: a.name, title: a.title, phone: a.phone, email: a.email, address: a.address })), settings: settings() };
  }
  if (d.action === 'updateApp') {
    const r = findRow('Applications', d.id); if (!r) return { ok: false, error: 'আবেদন পাওয়া যায়নি' };
    const s = sh('Applications'), old = s.getRange(r, 1, 1, 10).getValues()[0];
    const paid = Number(d.paid) || 0, oldPaid = Number(old[8]) || 0;
    if (paid < oldPaid) return { ok: false, error: 'জমা কমানো যাবে না। ভুল হলে "আয়" মেনু থেকে ওই এন্ট্রি মুছুন।' };
    s.getRange(r, 7, 1, 4).setValues([[tx(d.status, 30), Number(d.bill) || 0, paid, tx(d.note, 300)]]);
    if (paid > oldPaid) add('Accounts', { id: newId('E'), date: today(), type: 'আয়', desc: old[2] + ' - ' + old[3] + ' (' + d.id + ')', amount: paid - oldPaid, appId: d.id });
    return { ok: true };
  }
  if (d.action === 'addEntry') {
    add('Accounts', { id: newId('E'), date: tx(d.date, 10) || today(), type: d.type === 'ব্যয়' ? 'ব্যয়' : 'আয়', desc: tx(d.desc, 200), amount: amt(d.amount), appId: '' });
    return { ok: true };
  }
  if (d.action === 'addCash') {
    add('Cash', { id: newId('C'), date: tx(d.date, 10) || today(), type: d.type === 'উত্তোলন' ? 'উত্তোলন' : 'জমা', place: tx(d.place, 40), amount: amt(d.amount), note: tx(d.note, 200) });
    return { ok: true };
  }
  if (d.action === 'addStaff') {
    if (!tx(d.name, 100)) return { ok: false, error: 'নাম দিন' };
    add('Staff', { id: newId('S'), name: tx(d.name, 100), title: tx(d.title, 60), phone: tx(d.phone, 20), address: tx(d.address, 200), joined: tx(d.joined, 10) || today(), salary: Number(d.salary) || 0 });
    return { ok: true };
  }
  if (d.action === 'addSalary') {
    const st = rows('Staff').find(x => x.id === d.staffId); if (!st) return { ok: false, error: 'স্টাফ বেছে নিন' };
    add('Salaries', { id: newId('P'), date: tx(d.date, 10) || today(), staffId: st.id, staffName: st.name, month: tx(d.month, 7), amount: amt(d.amount), note: tx(d.note, 200) });
    return { ok: true };
  }
  if (d.action === 'addAdmin') {
    const name = tx(d.name, 100), phone = tx(d.phone, 20), pw = String(d.password || '');
    if (!name || !phone || pw.length < 6) return { ok: false, error: 'নাম, মোবাইল ও কমপক্ষে ৬ অক্ষরের পাসওয়ার্ড দিন' };
    if (rows('Admins').some(a => String(a.phone) === phone)) return { ok: false, error: 'এই মোবাইল নম্বরে অ্যাডমিন আছে' };
    const salt = Utilities.getUuid();
    add('Admins', { id: newId('M'), name, title: tx(d.title, 60), phone, email: tx(d.email, 100), address: tx(d.address, 200), hash: hash(pw, salt), salt });
    return { ok: true };
  }
  if (DEL[d.action]) {
    const n = DEL[d.action], r = findRow(n, d.id); if (!r) return { ok: true };
    if (n === 'Accounts') {
      const [, , type, , amount, appId] = sh(n).getRange(r, 1, 1, 6).getValues()[0];
      if (type === 'আয়' && appId) { const ar = findRow('Applications', appId); if (ar) { const c = sh('Applications').getRange(ar, 9); c.setValue(Math.max(0, (Number(c.getValue()) || 0) - Number(amount))); } }
    }
    sh(n).deleteRow(r); return { ok: true };
  }
  if (d.action === 'delAdmin') {
    if (d.id === who.id) return { ok: false, error: 'নিজেকে মুছা যাবে না' };
    const r = findRow('Admins', d.id); if (r) sh('Admins').deleteRow(r); return { ok: true };
  }
  if (d.action === 'saveSettings') {
    const s = sh('Settings');
    SET_KEYS.forEach(k => {
      if (d[k] === undefined) return;
      let v = String(d[k]).slice(0, k === 'logo' ? 45000 : 300);
      if (k === 'logo' && v && !/^data:image\//.test(v)) return;
      const r = findRow('Settings', k);
      if (r) s.getRange(r, 2).setNumberFormat('@').setValue(v); else add('Settings', { key: k, value: v });
    });
    return { ok: true };
  }
  return { ok: false, error: 'অজানা অনুরোধ' };
}
