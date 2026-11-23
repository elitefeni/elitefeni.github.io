// Google Sheet > Extensions > Apps Script এ এই কোড পেস্ট করুন (আগের কোড মুছে)।
// Script Properties এ ADMIN_PASS দিন (মূল অ্যাডমিন: আইডি owner)।
// কোড বদলানোর পর Deploy > Manage deployments > Edit > New version দিয়ে আবার Deploy করুন।
const SCH = {
  Applications: { h: ['আইডি','তারিখ','সেবা','নাম','ফোন','বিবরণ','স্ট্যাটাস','মোট বিল','জমা','নোট'], k: ['id','date','service','name','phone','details','status','bill','paid','note'] },
  Vouchers: { h: ['আইডি','ধরন','নম্বর','তারিখ','আইটেম','মোট','পরিশোধ','আবেদন আইডি'], k: ['id','type','no','date','items','total','paid','appId'] },
  CashEntries: { h: ['আইডি','তারিখ','ধরন','আইটেম','মোট'], k: ['id','date','type','items','total'] },
  Staff: { h: ['আইডি','নাম','পদবি','মোবাইল','ঠিকানা','যোগদান','মূল বেতন'], k: ['id','name','title','phone','address','joined','salary'] },
  Salaries: { h: ['আইডি','তারিখ','স্টাফ আইডি','স্টাফ','মাস','পরিমাণ','নোট'], k: ['id','date','staffId','staffName','month','amount','note'] },
  Admins: { h: ['আইডি','নাম','পদবি','মোবাইল','ইমেইল','ঠিকানা','হ্যাশ','সল্ট'], k: ['id','name','title','phone','email','address','hash','salt'] },
  Settings: { h: ['কী','মান'], k: ['key','value'] },
  Services: { h: ['আইডি','সেবার নাম','বিবরণ','ফিল্ড'], k: ['id','name','desc','fields'] },
  Packages: { h: ['আইডি','সেবা আইডি','নাম','অবস্থা','তথ্য'], k: ['id','serviceId','title','status','data'] }
};
const KINDS = { voucher: 'Vouchers', cash: 'CashEntries', staff: 'Staff', salary: 'Salaries', admin: 'Admins', service: 'Services', package: 'Packages', app: 'Applications' };
const JSONK = { items: 1, fields: 1, data: 1 };
const SET_KEYS = ['siteName','logo','address','phone','email','facebook','instagram','youtube','hours','hero'];
const SEED = [
  ['হজ্জ প্রসেসিং', 'নিবন্ধন, ভিসা, আবাসন ও ফ্লাইটসহ পূর্ণ প্রসেসিং এবং যাত্রার আগের নির্দেশনা।', ['খরচ','সময়কাল','আবাসন','ফ্লাইট','অন্যান্য সুবিধা']],
  ['ওমরাহ প্রসেসিং', 'ওমরাহ প্যাকেজ, ভিসা, হোটেল ও ফ্লাইটসহ সম্পূর্ণ সহযোগিতা।', ['খরচ','সময়কাল','মক্কার হোটেল','মদিনার হোটেল','ফ্লাইট','অন্যান্য সুবিধা']],
  ['পাসপোর্ট (নতুন/রিনিউ)', 'নতুন পাসপোর্ট, রিনিউ ও তথ্য সংশোধনের আবেদন, ফর্ম পূরণ ও ফি জমা।', ['ধরন','প্রসেসিং সময়','ফি','প্রয়োজনীয় কাগজ']],
  ['ভিসা প্রসেসিং', 'ভ্রমণ, কর্ম ও অন্যান্য ভিসার কাগজপত্র যাচাই ও আবেদন।', ['দেশ','ভিসার ধরন','সময়','ফি','প্রয়োজনীয় কাগজ']],
  ['টিকেট প্রসেসিং', 'দেশি ও আন্তর্জাতিক ফ্লাইটের টিকেট বুকিং, পরিবর্তন ও বাতিল।', ['রুট','এয়ারলাইন্স','ভাড়া','অন্যান্য তথ্য']],
  ['ল্যান্ড ও হাউজ প্রপার্টিজ', 'জমি ও বাড়ি কেনাবেচা, কাগজ যাচাই ও দলিল-সংক্রান্ত সহযোগিতা।', ['অবস্থান','আয়তন','মূল্য','অন্যান্য তথ্য']]
];

const dflt = k => k === 'data' ? {} : [];
function sh(n) {
  const ss = SpreadsheetApp.getActive(); let s = ss.getSheetByName(n);
  if (!s) { s = ss.insertSheet(n); s.appendRow(SCH[n].h); s.setFrozenRows(1); s.getRange(1, 1, 1, SCH[n].h.length).setFontWeight('bold'); }
  return s;
}
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const today = () => Utilities.formatDate(new Date(), 'Asia/Dhaka', 'yyyy-MM-dd');
const tx = (v, l) => String(v == null ? '' : v).trim().slice(0, l);
const pad = n => ('0000' + n).slice(-4);
const hash = (p, salt) => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + p));
const newId = p => p + Date.now().toString(36) + Math.floor(Math.random() * 1296).toString(36);
function cnt(p, peek) {
  const P = PropertiesService.getScriptProperties(), k = 'cnt_' + p, b = P.getProperty(k);
  let n = b === null ? (p === 'A' ? Math.max(0, sh('Applications').getLastRow() - 1) : 0) : Number(b);
  n++; if (!peek) P.setProperty(k, String(n)); return n;
}
function rows(n) {
  const v = sh(n).getDataRange().getValues(); v.shift();
  return v.map(r => {
    const o = {};
    SCH[n].k.forEach((k, j) => {
      let x = r[j];
      if (x instanceof Date) x = Utilities.formatDate(x, 'Asia/Dhaka', 'yyyy-MM-dd');
      if (k in JSONK) { try { x = JSON.parse(x); } catch (e) { x = dflt(k); } if (typeof x !== 'object' || x === null) x = dflt(k); }
      o[k] = x;
    });
    return o;
  });
}
function findRow(n, id) {
  const s = sh(n), ids = s.getRange(1, 1, s.getLastRow(), 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(id)); return i < 1 ? 0 : i + 1;
}
const cell = (k, v) => k in JSONK ? JSON.stringify(v == null ? dflt(k) : v) : (v == null ? '' : v);
function add(n, o) { const s = sh(n), r = s.getLastRow() + 1, k = SCH[n].k; s.getRange(r, 1, 1, k.length).setNumberFormat('@').setValues([k.map(x => cell(x, o[x]))]); }
function upd(n, r, o) { const k = SCH[n].k; sh(n).getRange(r, 1, 1, k.length).setNumberFormat('@').setValues([k.map(x => cell(x, o[x]))]); }
function settings() { const o = {}; rows('Settings').forEach(r => o[r.key] = r.value); return o; }
function seed() {
  const P = PropertiesService.getScriptProperties(); if (P.getProperty('seeded')) return;
  P.setProperty('seeded', '1');
  if (!rows('Services').length) SEED.forEach(([name, desc, fields], i) => add('Services', { id: newId('V' + i), name, desc, fields }));
}
const amtItems = (arr, label) => {
  const a = (Array.isArray(arr) ? arr : []).slice(0, 60).map(x => ({ t: tx(x.t, 200), a: Number(x.a) || 0 })).filter(x => x.a > 0);
  if (!a.length) throw new Error('কমপক্ষে একটি সারিতে পরিমাণ দিন'); return a;
};

const CLEAN = {
  voucher(r, old) {
    if (old && old.appId) throw new Error('আবেদন থেকে তৈরি রশিদ সম্পাদনা করা যায় না');
    const type = old ? old.type : (r.type === 'ব্যয়' ? 'ব্যয়' : 'আয়'), items = amtItems(r.items), total = items.reduce((t, x) => t + x.a, 0);
    return { type, no: old ? old.no : (type === 'আয়' ? 'R-' : 'V-') + pad(cnt(type === 'আয়' ? 'R' : 'V')), date: tx(r.date, 10) || today(), items, total,
      paid: type === 'ব্যয়' ? Math.min(total, Math.max(0, Number(r.paid) || 0)) : 0, appId: old ? old.appId : '' };
  },
  cash(r) { const items = amtItems(r.items); return { date: tx(r.date, 10) || today(), type: r.type === 'উত্তোলন' ? 'উত্তোলন' : 'জমা', items, total: items.reduce((t, x) => t + x.a, 0) }; },
  staff(r) {
    if (!tx(r.name, 100)) throw new Error('নাম দিন');
    return { name: tx(r.name, 100), title: tx(r.title, 60), phone: tx(r.phone, 20), address: tx(r.address, 200), joined: tx(r.joined, 10) || today(), salary: Number(r.salary) || 0 };
  },
  salary(r) {
    const st = rows('Staff').find(x => x.id === r.staffId); if (!st) throw new Error('স্টাফ বেছে নিন');
    const a = Number(r.amount); if (!(a > 0)) throw new Error('সঠিক পরিমাণ দিন');
    return { date: tx(r.date, 10) || today(), staffId: st.id, staffName: st.name, month: tx(r.month, 7), amount: a, note: tx(r.note, 200) };
  },
  admin(r, old) {
    const name = tx(r.name, 100), phone = tx(r.phone, 20), pw = String(r.password || '');
    if (!name || !phone) throw new Error('নাম ও মোবাইল নম্বর দিন');
    if (rows('Admins').some(a => String(a.phone) === phone && (!old || a.id !== old.id))) throw new Error('এই মোবাইল নম্বরে অ্যাডমিন আছে');
    let h = old ? old.hash : '', salt = old ? old.salt : '';
    if (pw || !old) { if (pw.length < 6) throw new Error('পাসওয়ার্ড কমপক্ষে ৬ অক্ষরের দিন'); salt = Utilities.getUuid(); h = hash(pw, salt); }
    return { name, title: tx(r.title, 60), phone, email: tx(r.email, 100), address: tx(r.address, 200), hash: h, salt };
  },
  service(r) {
    if (!tx(r.name, 100)) throw new Error('সেবার নাম দিন');
    const f = []; (Array.isArray(r.fields) ? r.fields : []).forEach(x => { x = tx(x, 60); if (x && f.indexOf(x) < 0) f.push(x); });
    return { name: tx(r.name, 100), desc: tx(r.desc, 300), fields: f.slice(0, 30) };
  },
  package(r) {
    if (!rows('Services').some(s => s.id === r.serviceId)) throw new Error('সেবা পাওয়া যায়নি');
    if (!tx(r.title, 150)) throw new Error('প্যাকেজ/আইটেমের নাম দিন');
    const data = {}; Object.keys(r.data || {}).slice(0, 30).forEach(k => data[tx(k, 60)] = tx(r.data[k], 500));
    return { serviceId: r.serviceId, title: tx(r.title, 150), status: r.status === 'বন্ধ' ? 'বন্ধ' : 'চালু', data };
  }
};

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
  if (d.action === 'public') {
    seed(); const s = settings(), o = {}; SET_KEYS.forEach(k => o[k] = s[k] || '');
    const pk = rows('Packages').filter(p => p.status === 'চালু');
    return { ok: true, settings: o, services: rows('Services').map(v => ({ id: v.id, name: v.name, desc: v.desc, fields: v.fields, packages: pk.filter(p => p.serviceId === v.id).map(p => ({ id: p.id, title: p.title, data: p.data })) })) };
  }
  if (d.action === 'apply') {
    const name = tx(d.name, 100), phone = tx(d.phone, 20), service = tx(d.service, 100);
    if (!name || !/^[0-9+\- ]{10,15}$/.test(phone)) return { ok: false, error: 'নাম ও সঠিক মোবাইল নম্বর দিন' };
    if (!service) return { ok: false, error: 'সেবা বেছে নিন' };
    const id = 'A' + Utilities.formatDate(new Date(), 'Asia/Dhaka', 'yyMMdd') + '-' + cnt('A');
    add('Applications', { id, date: today(), service, name, phone, details: tx(d.details, 800), status: 'নতুন', bill: 0, paid: 0, note: '' });
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
    seed();
    return { ok: true, me: who, apps: rows('Applications').reverse(), vouchers: rows('Vouchers').reverse(), cash: rows('CashEntries').reverse(), staff: rows('Staff'), salaries: rows('Salaries').reverse(),
      admins: rows('Admins').map(a => ({ id: a.id, name: a.name, title: a.title, phone: a.phone, email: a.email, address: a.address })),
      services: rows('Services'), packages: rows('Packages'), settings: settings(), next: { R: 'R-' + pad(cnt('R', true)), V: 'V-' + pad(cnt('V', true)) } };
  }
  if (d.action === 'updateApp') {
    const r = findRow('Applications', d.id); if (!r) return { ok: false, error: 'আবেদন পাওয়া যায়নি' };
    const s = sh('Applications'), old = s.getRange(r, 1, 1, 10).getValues()[0];
    const paid = Number(d.paid) || 0, oldPaid = Number(old[8]) || 0;
    if (paid < oldPaid) return { ok: false, error: 'জমা কমানো যাবে না। ভুল হলে "আয়" মেনু থেকে ওই রশিদ মুছুন।' };
    s.getRange(r, 7, 1, 4).setValues([[tx(d.status, 30), Number(d.bill) || 0, paid, tx(d.note, 300)]]);
    if (paid > oldPaid) add('Vouchers', { id: newId('R'), type: 'আয়', no: 'R-' + pad(cnt('R')), date: today(), items: [{ t: old[2] + ' - ' + old[3] + ' (' + d.id + ')', a: paid - oldPaid }], total: paid - oldPaid, paid: 0, appId: d.id });
    return { ok: true };
  }
  if (d.action === 'save') {
    const n = KINDS[d.kind]; if (!n || d.kind === 'app') return { ok: false, error: 'অজানা ধরন' };
    const rec = d.rec || {}; let r = 0, old = null;
    if (rec.id) { r = findRow(n, rec.id); if (!r) return { ok: false, error: 'রেকর্ড পাওয়া যায়নি' }; old = rows(n)[r - 2]; }
    const o = CLEAN[d.kind](rec, old);
    if (old) upd(n, r, Object.assign({}, old, o, { id: old.id })); else { o.id = newId(d.kind[0].toUpperCase()); add(n, o); }
    return { ok: true };
  }
  if (d.action === 'del') {
    const n = KINDS[d.kind]; if (!n) return { ok: false, error: 'অজানা ধরন' };
    if (d.kind === 'admin' && d.id === who.id) return { ok: false, error: 'নিজেকে মুছা যাবে না' };
    const r = findRow(n, d.id); if (!r) return { ok: true };
    if (d.kind === 'voucher') {
      const v = rows(n)[r - 2];
      if (v.type === 'আয়' && v.appId) { const ar = findRow('Applications', v.appId); if (ar) { const c = sh('Applications').getRange(ar, 9); c.setValue(Math.max(0, (Number(c.getValue()) || 0) - Number(v.total))); } }
    }
    if (d.kind === 'service') { const s = sh('Packages'), v = s.getDataRange().getValues(); for (let i = v.length - 1; i >= 1; i--) if (String(v[i][1]) === String(d.id)) s.deleteRow(i + 1); }
    sh(n).deleteRow(r); return { ok: true };
  }
  if (d.action === 'saveSettings') {
    const s = sh('Settings');
    SET_KEYS.forEach(k => {
      if (d[k] === undefined) return;
      const v = String(d[k]).slice(0, k === 'logo' || k === 'hero' ? 45000 : 300);
      if ((k === 'logo' || k === 'hero') && v && !/^data:image\//.test(v)) return;
      const r = findRow('Settings', k);
      if (r) s.getRange(r, 2).setNumberFormat('@').setValue(v); else add('Settings', { key: k, value: v });
    });
    return { ok: true };
  }
  return { ok: false, error: 'অজানা অনুরোধ' };
}
