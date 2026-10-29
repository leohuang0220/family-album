/**
 * 寶寶家庭相簿 — Google Apps Script 後端（零費用版）
 * 照片存在「你自己的 Google 雲端硬碟」，家人名單與權限存在同一個資料夾裡的 Google 試算表。
 * 所有權限檢查都在這裡執行：沒被邀請、或沒權限的人拿不到任何檔案。
 *
 * ① 只需要改下面兩個設定，然後在編輯器上方選 setup → 執行（只要一次）
 */
const OWNERS = ['leopupupu@gmail.com', 'li810606@gmail.com'];          // 爸爸、媽媽的 Gmail（永遠是管理者）
const CLIENT_ID = '757802579785-c6nk6pv185ctk99sb7og309gegud0smm.apps.googleusercontent.com'; // Google Cloud「OAuth 用戶端 ID」

/* ===================================================================== */

const TZ = 'Asia/Taipei';
const APP_URL = 'https://leohuang0220.github.io/family-album/';   // 通知信裡「打開相簿」的網址

// 第一次加上 Email 通知時，在編輯器選這個函式按「執行」一次，授權寄信
function authorizeMail() { Logger.log('授權完成，今天還能寄 ' + MailApp.getRemainingDailyQuota() + ' 封信'); }
const ROLES = ['admin', 'family', 'viewer'];
const VIS = ['everyone', 'family', 'parents'];
const TABLES = ['members', 'albums', 'photos', 'comments', 'activity', 'config', 'invites'];
const LOCKED = ['saveAlbum', 'deleteAlbum', 'saveMember', 'removeMember', 'setConfig', 'uploadStart', 'uploadFinish',
                'updatePhoto', 'like', 'addComment', 'delComment', 'deleteForever', 'setThumb', 'fixVideoThumbs', 'bulkState', 'saveSigs', 'setMyNotify', 'createInvite', 'revokeInvite'];
const CHUNK_MAX = 8 * 1024 * 1024;
const CACHE = CacheService.getScriptCache();

/* ---------- 一次性設定 ---------- */
function setup() {
  const p = PropertiesService.getScriptProperties();
  let root;
  if (p.getProperty('ROOT_ID')) root = DriveApp.getFolderById(p.getProperty('ROOT_ID'));
  else { root = DriveApp.createFolder('寶寶家庭相簿'); p.setProperty('ROOT_ID', root.getId()); }
  if (!p.getProperty('SHEET_ID')) {
    const ss = SpreadsheetApp.create('家庭相簿資料（請勿手動修改）');
    DriveApp.getFileById(ss.getId()).moveTo(root);
    TABLES.forEach(n => ss.getSheetByName(n) || ss.insertSheet(n));
    ss.getSheets().forEach(s => { if (TABLES.indexOf(s.getName()) < 0) ss.deleteSheet(s); });
    p.setProperty('SHEET_ID', ss.getId());
  }
  UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=test', { muteHttpExceptions: true }); // 觸發網路權限授權
  Logger.log('✅ 設定完成！照片資料夾：' + root.getUrl());
}

/* ---------- 入口 ---------- */
function doGet() {
  return ContentService.createTextOutput('寶寶家庭相簿後端運作中 ✅');
}

let _locked = false;
function doPost(e) {
  let out;
  const lock = LockService.getScriptLock();
  try {
    const req = JSON.parse(e.postData.contents);
    const who = req.session ? verifySession(req.session) : verify(req.token);
    let me = getMe(who.email, who.name);
    // 用邀請連結加入：還不是家人的人，憑邀請碼加入
    if (!me && req.action === 'joinInvite') { lock.waitLock(25000); _locked = true; me = joinInvite(req.code, who); }
    if (!me) out = { ok: false, error: req.session ? 'SESSION_INVALID' : 'NOT_INVITED', email: who.email };
    else {
      const fn = req.action === 'joinInvite' ? (() => ({ joined: true })) : API[req.action];
      if (!fn) throw err('BAD_ACTION');
      if (LOCKED.indexOf(req.action) >= 0) { lock.waitLock(25000); _locked = true; }
      out = { ok: true, data: fn(req, me) };
      // 發一張長效的登入憑證，手機（尤其 iPhone）就不用每次重新登入 Google
      if (!req.session || req.action === 'bootstrap') out.session = makeSession(who);
    }
  } catch (x) {
    console.error(x);
    out = { ok: false, error: x.code || String(x.message || x) };
  } finally {
    if (_locked) { lock.releaseLock(); _locked = false; }
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function err(code) { const e = new Error(code); e.code = code; return e; }
function need(cond, code) { if (!cond) throw err(code || 'FORBIDDEN'); }
function uid(prefix) { return prefix + Date.now().toString(36) + Utilities.getUuid().replace(/-/g, '').slice(0, 8); }

/* ---------- 驗證 Google 登入 ---------- */
function verify(token) {
  need(token, 'NO_TOKEN');
  const key = 'tok:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token)).slice(0, 43);
  const hit = CACHE.get(key);
  if (hit) return JSON.parse(hit);
  const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token), { muteHttpExceptions: true });
  need(r.getResponseCode() === 200, 'TOKEN_INVALID');
  const p = JSON.parse(r.getContentText());
  need(p.aud === CLIENT_ID && String(p.email_verified) === 'true' && /accounts\.google\.com$/.test(p.iss || ''), 'TOKEN_INVALID');
  const who = { email: String(p.email).toLowerCase(), name: p.name || '' };
  const ttl = Math.min(21600, Math.max(30, Number(p.exp) - Math.floor(Date.now() / 1000) - 30));
  CACHE.put(key, JSON.stringify(who), ttl);
  return who;
}

/* ---------- 保持登入：自己簽發的登入憑證（90 天，每次打開 App 自動延長） ---------- */
const SESSION_DAYS = 90;
function sessionSecret() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('SESSION_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SESSION_SECRET', s); }
  return s;
}
function signSession(body) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(body, sessionSecret())); }
function makeSession(who) {
  const body = Utilities.base64EncodeWebSafe(JSON.stringify({ e: who.email, n: who.name || '', x: Date.now() + SESSION_DAYS * 864e5 }));
  return 's1.' + body + '.' + signSession(body);
}
function verifySession(s) {
  const parts = String(s || '').split('.');
  need(parts.length === 3 && parts[0] === 's1' && signSession(parts[1]) === parts[2], 'SESSION_INVALID');
  let o; try { o = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[1])).getDataAsString()); } catch (x) { throw err('SESSION_INVALID'); }
  need(o && o.e && Number(o.x) > Date.now(), 'SESSION_INVALID');
  return { email: String(o.e).toLowerCase(), name: o.n || '' };
}

/* ---------- 試算表資料表（含快取） ---------- */
function prop(k) { const v = PropertiesService.getScriptProperties().getProperty(k); if (!v) throw err('NOT_SETUP'); return v; }
let _ss = null;
function ss() { return _ss || (_ss = SpreadsheetApp.openById(prop('SHEET_ID'))); }
function sheet(n) { return ss().getSheetByName(n) || ss().insertSheet(n); }
function readAll(n) {
  const ver = CACHE.get('v:' + n);
  if (ver) {
    const m = JSON.parse(ver), keys = [];
    for (let i = 0; i < m.c; i++) keys.push('t:' + n + ':' + m.id + ':' + i);
    const got = CACHE.getAll(keys);
    if (Object.keys(got).length === m.c) return JSON.parse(keys.map(k => got[k]).join(''));
  }
  const s = sheet(n), last = s.getLastRow();
  const rows = last ? s.getRange(1, 1, last, 2).getValues().filter(r => r[0] !== '' && r[1] !== '').map(r => JSON.parse(r[1])) : [];
  try {
    const str = JSON.stringify(rows), id = Utilities.getUuid().slice(0, 8), parts = {};
    let c = 0;
    for (let i = 0; i < str.length; i += 25000) parts['t:' + n + ':' + id + ':' + (c++)] = str.slice(i, i + 25000);
    if (c <= 200) { CACHE.putAll(parts, 900); CACHE.put('v:' + n, JSON.stringify({ id: id, c: c }), 900); }
  } catch (x) {}
  return rows;
}
function invalidate(n) { CACHE.remove('v:' + n); }
function rowOf(s, id) {
  const last = s.getLastRow(); if (!last) return -1;
  const ids = s.getRange(1, 1, last, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 1;
  return -1;
}
function put(n, obj) {
  const s = sheet(n), i = rowOf(s, obj.id), row = [String(obj.id), JSON.stringify(obj)];
  if (i > 0) s.getRange(i, 1, 1, 2).setValues([row]); else s.appendRow(row);
  invalidate(n);
}
// 一次改很多筆（只讀一次 id 欄，比一筆一筆找快很多）
function putMany(n, objs) {
  if (!objs.length) return;
  const s = sheet(n), last = s.getLastRow(), ids = last ? s.getRange(1, 1, last, 1).getValues().map(r => String(r[0])) : [];
  objs.forEach(o => { const i = ids.indexOf(String(o.id)), row = [String(o.id), JSON.stringify(o)]; if (i >= 0) s.getRange(i + 1, 1, 1, 2).setValues([row]); else { s.appendRow(row); ids.push(String(o.id)); } });
  invalidate(n);
}
function del(n, id) { const s = sheet(n), i = rowOf(s, id); if (i > 0) s.deleteRow(i); invalidate(n); }
function getOne(n, id) { return readAll(n).find(x => String(x.id) === String(id)) || null; }
function log(me, text) { sheet('activity').appendRow([uid('l'), JSON.stringify({ at: Date.now(), actor: me.email, actorName: me.name || me.email, text: String(text).slice(0, 200) })]); }

/* ---------- 邀請連結（不用先知道對方的 Gmail） ---------- */
function inviteOk(inv) { return inv && !inv.revoked && (!inv.expiresAt || inv.expiresAt > Date.now()) && (!inv.maxUses || (inv.uses || 0) < inv.maxUses); }
function joinInvite(code, who) {
  const inv = getOne('invites', String(code || ''));
  need(inv && !inv.revoked, 'INVITE_INVALID');
  need(!inv.expiresAt || inv.expiresAt > Date.now(), 'INVITE_EXPIRED');
  need(!inv.maxUses || (inv.uses || 0) < inv.maxUses, 'INVITE_USED_UP');
  const m = { id: who.email, email: who.email, name: String(inv.name || who.name || who.email.split('@')[0]).slice(0, 30), relation: inv.relation || '',
    role: ROLES.indexOf(inv.role) >= 0 && inv.role !== 'admin' ? inv.role : 'viewer', canDownload: inv.canDownload === true, notify: true,
    addedAt: Date.now(), via: 'invite', inviteId: inv.id, invitedBy: inv.createdBy || '' };
  put('members', m);
  inv.uses = (inv.uses || 0) + 1; inv.joined = (inv.joined || []).concat([who.email]); put('invites', inv);
  try { log(m, '透過邀請連結加入了相簿'); } catch (x) {}
  try {   // 通知爸媽有人加入
    const baby = (getOne('config', 'baby') || {}).babyName || '寶寶';
    OWNERS.forEach(o => MailApp.sendEmail({ to: o, subject: m.name + ' 加入了' + baby + '的相簿', name: baby + '的相簿',
      htmlBody: '<p><b>' + escHtml(m.name) + '</b>（' + escHtml(m.email) + '）用邀請連結加入了相簿。</p><p>如果不認識這個人，打開相簿 →「管理」→ 找到他按「編輯」→「移除」。</p><p><a href="' + APP_URL + '">打開相簿</a></p>' }));
  } catch (x) {}
  return getMe(who.email, who.name);
}

/* ---------- 權限 ---------- */
function getMe(email, name) {
  const owner = OWNERS.map(s => s.toLowerCase()).indexOf(email) >= 0;
  let m = getOne('members', email);
  if (owner && !m) {
    m = { id: email, email: email, name: name || '爸媽', relation: '', role: 'admin', canDownload: true, owner: true, addedAt: Date.now() };
    put('members', m);
  }
  if (!m) return null;
  const me = Object.assign({}, m, { email: email });
  if (owner) Object.assign(me, { role: 'admin', canDownload: true, owner: true });
  return me;
}
function isAdmin(me) { return me.role === 'admin'; }
function canSee(a, me) {
  return !!a && (isAdmin(me) || a.visibility === 'everyone' || (a.visibility === 'family' && me.role === 'family') || (a.allowed || []).indexOf(me.email) >= 0);
}
function canUploadTo(a, me) { return !!a && (isAdmin(me) || (me.role === 'family' && a.allowUpload === true && canSee(a, me))); }
function photoFor(id, me, allowHidden) {
  const p = getOne('photos', id); need(p, 'NOT_FOUND');
  const a = getOne('albums', p.albumId);
  need(canSee(a, me), 'FORBIDDEN');
  need(p.state === 'normal' || isAdmin(me) || allowHidden, 'FORBIDDEN');
  return p;
}
function publicPhoto(p) { const o = Object.assign({}, p); delete o.files; delete o.sig; delete o.sr; return o; }
function escHtml(s) { return String(s || '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[ch]); }
function validSig(s) { return (typeof s === 'string' && s.length <= 800 && /^[A-Za-z0-9+\/=]+$/.test(s)) ? s : ''; }

/* ---------- Drive ---------- */
function root() { return DriveApp.getFolderById(prop('ROOT_ID')); }
function albumFolder(a) {
  if (a.folderId) { try { const f = DriveApp.getFolderById(a.folderId); if (!f.isTrashed()) return f; } catch (x) {} }
  const f = root().createFolder(a.name || '相簿');
  a.folderId = f.getId(); put('albums', a);
  return f;
}
function driveFetch(url, opt) {
  opt = opt || {};
  opt.headers = Object.assign({ Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, opt.headers || {});
  opt.muteHttpExceptions = true;
  return UrlFetchApp.fetch(url, opt);
}
function validPh(v) { return (typeof v === 'string' && v.length < 3000 && /^data:image\/jpeg;base64,[A-Za-z0-9+\/=]+$/.test(v)) ? v : ''; }
function concatBytes() { let out = []; for (let i = 0; i < arguments.length; i++) out = out.concat(arguments[i]); return out; }
// 一次把多個小圖存進雲端硬碟（平行處理，比一個一個存快很多）
function multipartReq(name, folderId, bytes, tok) {
  const b = 'bb' + Utilities.getUuid().replace(/-/g, '');
  const head = Utilities.newBlob('--' + b + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + JSON.stringify({ name: name, parents: [folderId], mimeType: 'image/jpeg' }) + '\r\n--' + b + '\r\nContent-Type: image/jpeg\r\n\r\n').getBytes();
  const tail = Utilities.newBlob('\r\n--' + b + '--').getBytes();
  return { url: 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id', method: 'post', contentType: 'multipart/related; boundary=' + b,
    payload: concatBytes(head, bytes, tail), headers: { Authorization: 'Bearer ' + tok }, muteHttpExceptions: true };
}
function trashFiles(p) { Object.keys(p.files || {}).forEach(k => { try { DriveApp.getFileById(p.files[k]).setTrashed(true); } catch (x) {} }); }

/* ---------- API ---------- */
const API = {
  bootstrap(req, me) {
    if (!me.lastSeen || Date.now() - me.lastSeen > 10 * 60000) {
      const l = LockService.getScriptLock();
      if (l.tryLock(5000)) { try { const m = getOne('members', me.email); if (m) { m.lastSeen = Date.now(); put('members', m); } } finally { l.releaseLock(); } }
    }
    if (isAdmin(me)) purgeOldTrash(me);
    const albums = readAll('albums').filter(a => canSee(a, me));
    const ids = {}; albums.forEach(a => ids[a.id] = 1);
    const photos = readAll('photos').filter(p => ids[p.albumId] && (p.state === 'normal' || isAdmin(me))).map(publicPhoto);
    const cfg0 = getOne('config', 'baby') || {};
    const showSeen = cfg0.showPresence !== false;
    const members = readAll('members').map(m => isAdmin(me) ? m : (showSeen ? { email: m.email, name: m.name, relation: m.relation, role: m.role, lastSeen: m.lastSeen || 0 } : { email: m.email, name: m.name, relation: m.relation, role: m.role }));
    return { me: me, cfg: cfg0, albums: albums, photos: photos, members: members };
  },

  // 縮圖、顯示用圖片：一次取多張
  media(req, me) {
    const which = req.which === 'disp' ? 'disp' : 'thumb';
    const all = {}; readAll('photos').forEach(p => all[p.id] = p);
    const albums = {}; readAll('albums').forEach(a => albums[a.id] = a);
    const out = {}, want = (req.ids || []).slice(0, 40), fids = {};
    want.forEach(id => {
      const p = all[id];
      if (!p || !canSee(albums[p.albumId], me) || (p.state !== 'normal' && !isAdmin(me))) return;
      const fid = (p.files || {})[which] || (p.files || {}).thumb;
      if (fid) fids[id] = fid;
    });
    const cached = CACHE.getAll(Object.keys(fids).map(id => 'm:' + fids[id]));
    const miss = [];
    Object.keys(fids).forEach(id => { const k = 'm:' + fids[id]; if (cached[k]) out[id] = cached[k]; else miss.push(id); });
    if (miss.length) {
      const tok = ScriptApp.getOAuthToken();
      const rs = UrlFetchApp.fetchAll(miss.map(id => ({ url: 'https://www.googleapis.com/drive/v3/files/' + fids[id] + '?alt=media', headers: { Authorization: 'Bearer ' + tok }, muteHttpExceptions: true })));
      const toCache = {};
      rs.forEach((r, i) => {
        const id = miss[i]; if (r.getResponseCode() !== 200) return;
        const b64 = Utilities.base64Encode(r.getBlob().getBytes());
        out[id] = b64; if (b64.length < 95000) toCache['m:' + fids[id]] = b64;
      });
      try { if (Object.keys(toCache).length) CACHE.putAll(toCache, 21600); } catch (x) {}
    }
    return out;
  },

  // 原檔分段下載（影片播放、下載原圖）
  chunk(req, me) {
    const p = photoFor(req.id, me);
    need(p.kind === 'video' || me.canDownload, 'NO_DOWNLOAD');
    const fid = p.files.full; need(fid, 'NOT_FOUND');
    const start = Math.max(0, Number(req.start) || 0), len = Math.min(CHUNK_MAX, Number(req.len) || CHUNK_MAX);
    const r = driveFetch('https://www.googleapis.com/drive/v3/files/' + fid + '?alt=media', { headers: { Range: 'bytes=' + start + '-' + (start + len - 1) } });
    need(r.getResponseCode() === 206 || r.getResponseCode() === 200, 'DRIVE_' + r.getResponseCode());
    const cr = r.getAllHeaders()['Content-Range'] || '';
    const total = Number((cr.split('/')[1]) || (p.sizes || {}).full || 0);
    return { b64: Utilities.base64Encode(r.getBlob().getBytes()), total: total };
  },

  uploadStart(req, me) {
    const a = getOne('albums', req.albumId);
    need(canUploadTo(a, me), 'NO_UPLOAD');
    const o = req.orig || {};
    need(o.size > 0 && o.size <= 300 * 1024 * 1024, 'TOO_BIG');
    need(/^(image|video)\//.test(o.mime || ''), 'BAD_TYPE');
    const id = uid('p'), folder = albumFolder(a);
    const meta = req.meta || {};
    const stamp = Utilities.formatDate(new Date(Number(meta.takenAt) || Date.now()), TZ, 'yyyy-MM-dd_HHmm');
    const files = {}, sizes = {}, tok = ScriptApp.getOAuthToken(), reqs = [], kinds = [];
    ['thumb', 'disp'].forEach(k => {
      if (!req[k]) return;
      const bytes = Utilities.base64Decode(req[k]);
      need(bytes.length < 6 * 1024 * 1024, 'TOO_BIG');
      sizes[k] = bytes.length; kinds.push(k);
      reqs.push(multipartReq(stamp + '_' + id + '_' + k + '.jpg', folder.getId(), bytes, tok));
    });
    const ext = (String(o.name || '').split('.').pop() || 'bin').toLowerCase().replace(/[^a-z0-9]/g, '').slice(0, 5) || 'bin';
    reqs.push({ url: 'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&fields=id', method: 'post', contentType: 'application/json; charset=UTF-8',
      headers: { Authorization: 'Bearer ' + tok, 'X-Upload-Content-Type': o.mime, 'X-Upload-Content-Length': String(o.size) },
      payload: JSON.stringify({ name: stamp + '_' + id + '.' + ext, parents: [folder.getId()] }), muteHttpExceptions: true });
    const rs = UrlFetchApp.fetchAll(reqs);
    kinds.forEach((k, i) => { need(rs[i].getResponseCode() === 200, 'DRIVE_' + rs[i].getResponseCode()); files[k] = JSON.parse(rs[i].getContentText()).id; });
    const r = rs[rs.length - 1];
    need(r.getResponseCode() === 200, 'DRIVE_' + r.getResponseCode());
    const h = r.getAllHeaders(), session = h.Location || h.location;
    need(session, 'DRIVE_SESSION');
    const rec = { id: id, albumId: a.id, uploader: me.email, session: session, size: o.size, mime: o.mime, files: files, sizes: sizes, meta: meta };
    CACHE.put('up:' + id, JSON.stringify(rec), 21600);
    return { uploadId: id };
  },

  uploadChunk(req, me) {
    const raw = CACHE.get('up:' + req.uploadId); need(raw, 'UPLOAD_EXPIRED');
    const rec = JSON.parse(raw); need(rec.uploader === me.email, 'FORBIDDEN');
    const bytes = Utilities.base64Decode(req.b64), start = Number(req.start);
    need(bytes.length > 0 && bytes.length <= CHUNK_MAX, 'BAD_CHUNK');
    const end = start + bytes.length - 1;
    const r = UrlFetchApp.fetch(rec.session, {
      method: 'put', payload: Utilities.newBlob(bytes, rec.mime), contentType: rec.mime,
      headers: { 'Content-Range': 'bytes ' + start + '-' + end + '/' + rec.size }, muteHttpExceptions: true,
    });
    const code = r.getResponseCode();
    if (code === 308) return { done: false };
    need(code === 200 || code === 201, 'DRIVE_' + code);
    rec.files.full = JSON.parse(r.getContentText()).id; rec.sizes.full = rec.size;
    CACHE.put('up:' + req.uploadId, JSON.stringify(rec), 21600);
    return { done: true };
  },

  uploadFinish(req, me) {
    const raw = CACHE.get('up:' + req.uploadId); need(raw, 'UPLOAD_EXPIRED');
    const rec = JSON.parse(raw); need(rec.uploader === me.email && rec.files.full, 'UPLOAD_INCOMPLETE');
    const m = rec.meta || {};
    const p = {
      id: rec.id, albumId: rec.albumId, uploader: me.email, createdAt: Date.now(), state: 'normal', likes: [], commentCount: 0,
      kind: m.kind === 'video' ? 'video' : 'image', mime: rec.mime, takenAt: Number(m.takenAt) || Date.now(),
      w: Number(m.w) || 0, h: Number(m.h) || 0, duration: Number(m.duration) || 0,
      caption: String(m.caption || '').slice(0, 300), milestone: String(m.milestone || '').slice(0, 40),
      srcName: String(m.srcName || '').slice(0, 120), srcSize: Number(m.srcSize) || 0, ph: validPh(m.ph), thumbOk: m.thumbOk === true, sig: validSig(m.sig), sr: Number(m.sr) || 0,
      files: rec.files, sizes: rec.sizes, bytes: Object.keys(rec.sizes).reduce((s, k) => s + rec.sizes[k], 0),
    };
    put('photos', p);
    CACHE.remove('up:' + req.uploadId);
    return publicPhoto(p);
  },

  // 舊影片沒有縮圖：直接用 Google 雲端硬碟自己產生的影片縮圖補上（不用下載影片）
  fixVideoThumbs(req, me) {
    need(isAdmin(me));
    const bad = p => p.kind === 'video' && !p.thumbOk && p.state !== 'trash' && p.files && p.files.full && p.files.thumb;
    const todo = readAll('photos').filter(bad).slice(0, 12);
    if (!todo.length) return { fixed: 0, left: 0, photos: [] };
    const tok = ScriptApp.getOAuthToken(), auth = { Authorization: 'Bearer ' + tok };
    const metas = UrlFetchApp.fetchAll(todo.map(p => ({ url: 'https://www.googleapis.com/drive/v3/files/' + p.files.full + '?fields=thumbnailLink', headers: auth, muteHttpExceptions: true })));
    const links = metas.map(r => { try { return r.getResponseCode() === 200 ? (JSON.parse(r.getContentText()).thumbnailLink || '') : ''; } catch (x) { return ''; } });
    const want = []; links.forEach((u, i) => { if (u) want.push(i); });
    const imgs = want.length ? UrlFetchApp.fetchAll(want.map(i => ({ url: links[i].replace(/=s\d+$/, '') + '=s640', headers: auth, muteHttpExceptions: true }))) : [];
    const bytesFor = {};
    imgs.forEach((r, k) => { if (r.getResponseCode() === 200) { const b = r.getBlob().getBytes(); if (b.length > 800) bytesFor[want[k]] = b; } });
    todo.forEach((p, i) => {   // 拿不到連結的，改用 Apps Script 內建的縮圖
      if (bytesFor[i] || !links[i]) return;
      try { const t = DriveApp.getFileById(p.files.full).getThumbnail(); if (t) { const b = t.getBytes(); if (b.length > 800) bytesFor[i] = b; } } catch (x) {}
    });
    const idx = Object.keys(bytesFor).map(Number);
    const rs = idx.length ? UrlFetchApp.fetchAll(idx.map(i => ({ url: 'https://www.googleapis.com/upload/drive/v3/files/' + todo[i].files.thumb + '?uploadType=media', method: 'patch', contentType: 'image/jpeg', payload: bytesFor[i], headers: auth, muteHttpExceptions: true }))) : [];
    const done = [];
    rs.forEach((r, k) => {
      if (r.getResponseCode() !== 200) return;
      const p = todo[idx[k]]; p.thumbOk = true; p.tv = Date.now(); p.ph = '';
      put('photos', p); try { CACHE.remove('m:' + p.files.thumb); } catch (x) {}
      done.push(publicPhoto(p));
    });
    const left = readAll('photos').filter(bad).length;
    return { fixed: done.length, left: left, photos: done };
  },
  // 有人上傳新照片：寄一封 Email 給看得到這些照片、而且沒有關掉通知的家人
  notifyNew(req, me) {
    const want = {}; (req.ids || []).slice(0, 200).forEach(i => want[String(i)] = 1);
    const ps = readAll('photos').filter(p => want[p.id] && p.uploader === me.email && p.state === 'normal' && Date.now() - (p.createdAt || 0) < 6 * 3600e3)
      .sort((a, b) => b.takenAt - a.takenAt);
    if (!ps.length) return { sent: 0 };
    const albums = {}; readAll('albums').forEach(a => albums[a.id] = a);
    const baby = (getOne('config', 'baby') || {}).babyName || '寶寶', who = me.name || me.email.split('@')[0];
    const blobs = {}, thumbOf = p => { if (!(p.id in blobs)) { try { blobs[p.id] = DriveApp.getFileById(p.files.thumb).getBlob().setName(p.id + '.jpg'); } catch (x) { blobs[p.id] = null; } } return blobs[p.id]; };
    let quota = 0; try { quota = MailApp.getRemainingDailyQuota(); } catch (x) { throw err('MAIL_NOT_AUTHORIZED'); }
    let sent = 0;
    // 爸媽就算還沒登入過相簿（不在家人名單裡）也要收到
    const list = readAll('members').slice(), seen = {}; list.forEach(m => seen[m.email] = 1);
    OWNERS.map(e => e.toLowerCase()).forEach(e => { if (!seen[e]) list.push({ email: e, name: '' }); });
    list.forEach(m => {
      if (sent >= quota || m.email === me.email || m.notify === false) return;
      const viewer = getMe(m.email, m.name); if (!viewer) return;
      const vis = ps.filter(p => canSee(albums[p.albumId], viewer)); if (!vis.length) return;
      const vids = vis.filter(p => p.kind === 'video').length, imgs = vis.length - vids;
      const what = [imgs ? imgs + ' 張照片' : '', vids ? vids + ' 支影片' : ''].filter(Boolean).join('、');
      const names = [...new Set(vis.map(p => (albums[p.albumId] || {}).name).filter(Boolean))].map(n => '「' + escHtml(n) + '」').join('、');
      const inline = {}, cells = vis.slice(0, 4).map((p, i) => { const b = thumbOf(p); if (!b) return ''; inline['t' + i] = b;
        return '<td style="padding:3px"><img src="cid:t' + i + '" width="130" height="130" style="display:block;width:130px;height:130px;object-fit:cover;border-radius:10px"></td>'; }).join('');
      const link = APP_URL + '?p=' + encodeURIComponent(vis[0].albumId + '~' + vis[0].id);
      const html = '<div style="font-family:-apple-system,\'PingFang TC\',\'Microsoft JhengHei\',sans-serif;color:#1D2A35;max-width:560px">' +
        '<p style="font-size:17px;margin:0 0 6px"><b>' + escHtml(who) + '</b> 剛剛上傳了 ' + what + '</p>' +
        '<p style="color:#667785;margin:0 0 14px">放在 ' + names + (vis.length > 4 ? '，以下是其中 4 張' : '') + '</p>' +
        '<table cellspacing="0" cellpadding="0"><tr>' + cells + '</tr></table>' +
        '<p style="margin:20px 0"><a href="' + link + '" style="background:#2E7DA8;color:#fff;text-decoration:none;padding:12px 22px;border-radius:12px;font-weight:bold;display:inline-block">打開相簿看看</a></p>' +
        '<p style="color:#8A99A5;font-size:12px">不想收到這種信：<a href="' + APP_URL + '?notify=off" style="color:#8A99A5">點這裡關閉通知</a>。之後想再收到，打開相簿 →「我的」→ 打開「有新照片時寄 Email 通知我」。</p></div>';
      try {
        MailApp.sendEmail({ to: m.email, subject: baby + '的相簿新增了 ' + what, htmlBody: html, inlineImages: inline, name: baby + '的相簿' });
        sent++;
      } catch (x) { console.error('mail', m.email, x); }
    });
    return { sent: sent };
  },
  createInvite(req, me) {
    need(isAdmin(me));
    const x = req.invite || {}, days = Number(x.days) || 0;
    const inv = { id: Utilities.getUuid().replace(/-/g, '').slice(0, 16), role: ['family', 'viewer'].indexOf(x.role) >= 0 ? x.role : 'viewer',
      canDownload: x.canDownload === true, name: String(x.name || '').slice(0, 30), relation: String(x.relation || '').slice(0, 10),
      maxUses: x.multi ? 0 : 1, expiresAt: days > 0 ? Date.now() + days * 864e5 : 0, uses: 0, joined: [], createdBy: me.email, createdAt: Date.now() };
    put('invites', inv);
    return inv;
  },
  listInvites(req, me) {
    need(isAdmin(me));
    return readAll('invites').filter(inviteOk).sort((a, b) => b.createdAt - a.createdAt);
  },
  revokeInvite(req, me) {
    need(isAdmin(me));
    const inv = getOne('invites', String(req.id || '')); need(inv, 'NOT_FOUND');
    inv.revoked = true; put('invites', inv);
    return { ok: true };
  },
  setMyNotify(req, me) {
    const m = getOne('members', me.email); need(m, 'NOT_FOUND');
    m.notify = req.on === true; put('members', m);
    return { notify: m.notify };
  },
  testMail(req, me) {
    need(isAdmin(me));
    let quota; try { quota = MailApp.getRemainingDailyQuota(); } catch (x) { throw err('MAIL_NOT_AUTHORIZED'); }
    MailApp.sendEmail({ to: me.email, subject: '相簿 Email 通知測試', htmlBody: '<p>收到這封信，代表相簿的 Email 通知可以正常寄出。</p>', name: '寶寶家庭相簿' });
    return { quota: quota - 1 };
  },
  // 找重複照片用的畫面指紋：上傳時算好存起來，之後比對不用再下載縮圖
  sigs(req, me) {
    need(isAdmin(me));
    const out = {}; readAll('photos').forEach(p => { if (p.sig) out[p.id] = { s: p.sig, r: p.sr || 0 }; });
    return out;
  },
  saveSigs(req, me) {
    need(isAdmin(me));
    const want = {}; (req.items || []).slice(0, 500).forEach(x => { const s = validSig(x && x.s); if (s) want[String(x.id)] = { s: s, r: Number(x.r) || 0 }; });
    const changed = [];
    readAll('photos').forEach(p => { const w = want[p.id]; if (w && p.sig !== w.s) { p.sig = w.s; p.sr = w.r; changed.push(p); } });
    putMany('photos', changed);
    return { saved: changed.length };
  },
  // 多選刪除／隱藏／還原
  bulkState(req, me) {
    const admin = isAdmin(me), state = req.state;
    need((admin ? ['normal', 'hidden', 'trash'] : ['trash']).indexOf(state) >= 0, 'FORBIDDEN');
    const ids = (req.ids || []).slice(0, 300).map(String), want = {}; ids.forEach(i => want[i] = 1);
    const albums = {}; readAll('albums').forEach(a => albums[a.id] = a);
    const changed = [];
    readAll('photos').forEach(p => {
      if (!want[p.id] || !canSee(albums[p.albumId], me)) return;
      if (!admin && !(p.uploader === me.email && p.state === 'normal')) return;
      p.state = state;
      if (state === 'trash') p.trashedAt = Date.now(); else delete p.trashedAt;
      changed.push(p);
    });
    putMany('photos', changed);
    return { ids: changed.map(p => p.id) };
  },
  // 換掉影片縮圖（舊影片當初沒抓到畫面時用）
  setThumb(req, me) {
    const p = photoFor(req.id, me);
    need(isAdmin(me) || p.uploader === me.email, 'FORBIDDEN');
    need(p.files && p.files.thumb, 'NO_MEDIA');
    const bytes = Utilities.base64Decode(String(req.b64 || ''));
    need(bytes.length > 0 && bytes.length < 2 * 1024 * 1024, 'TOO_BIG');
    const r = driveFetch('https://www.googleapis.com/upload/drive/v3/files/' + p.files.thumb + '?uploadType=media', { method: 'patch', contentType: 'image/jpeg', payload: bytes });
    need(r.getResponseCode() === 200, 'DRIVE_' + r.getResponseCode());
    try { CACHE.remove('m:' + p.files.thumb); } catch (x) {}
    if (req.ph) p.ph = validPh(req.ph);
    p.thumbOk = true; p.tv = Date.now();
    put('photos', p);
    return publicPhoto(p);
  },
  updatePhoto(req, me) {
    const p = photoFor(req.id, me);
    const patch = req.patch || {}, admin = isAdmin(me), mine = p.uploader === me.email && p.state === 'normal';
    need(admin || mine, 'FORBIDDEN');
    if ('caption' in patch) p.caption = String(patch.caption || '').slice(0, 300);
    if ('milestone' in patch) p.milestone = String(patch.milestone || '').slice(0, 40);
    if ('ph' in patch) p.ph = validPh(patch.ph);
    if ('state' in patch) {
      const ok = admin ? ['normal', 'hidden', 'trash'] : ['normal', 'trash'];
      need(ok.indexOf(patch.state) >= 0, 'FORBIDDEN');
      p.state = patch.state;
      if (p.state === 'trash') p.trashedAt = Date.now(); else delete p.trashedAt;
    }
    put('photos', p);
    return publicPhoto(p);
  },

  like(req, me) {
    const p = photoFor(req.id, me);
    p.likes = (p.likes || []).filter(e => e !== me.email);
    if (req.on) p.likes.push(me.email);
    put('photos', p);
    return p.likes;
  },

  comments(req, me) {
    photoFor(req.id, me);
    return readAll('comments').filter(c => c.photoId === req.id).sort((a, b) => a.at - b.at);
  },
  addComment(req, me) {
    const p = photoFor(req.id, me);
    const text = String(req.text || '').trim().slice(0, 500); need(text, 'EMPTY');
    const c = { id: uid('c'), photoId: p.id, author: me.email, name: me.name || '', text: text, at: Date.now() };
    put('comments', c);
    p.commentCount = (p.commentCount || 0) + 1; put('photos', p);
    return c;
  },
  delComment(req, me) {
    const c = getOne('comments', req.cid); need(c, 'NOT_FOUND');
    need(isAdmin(me) || c.author === me.email, 'FORBIDDEN');
    del('comments', c.id);
    const p = getOne('photos', c.photoId);
    if (p) { p.commentCount = Math.max(0, (p.commentCount || 1) - 1); put('photos', p); }
    return true;
  },

  saveAlbum(req, me) {
    need(isAdmin(me));
    const x = req.album || {};
    let a = x.id ? getOne('albums', x.id) : null;
    if (x.id) need(a, 'NOT_FOUND');
    if (!a) a = { id: uid('a'), createdAt: Date.now(), name: '新相簿', visibility: 'everyone', allowed: [], allowUpload: false };
    if ('name' in x) a.name = String(x.name || '').trim().slice(0, 40) || a.name;
    if ('visibility' in x) { need(VIS.indexOf(x.visibility) >= 0, 'BAD_VIS'); a.visibility = x.visibility; }
    if ('allowed' in x) a.allowed = (x.allowed || []).map(e => String(e).toLowerCase()).slice(0, 50);
    if ('allowUpload' in x) a.allowUpload = x.allowUpload === true;
    if ('coverId' in x) a.coverId = String(x.coverId || '');
    put('albums', a);
    try { albumFolder(a).setName(a.name); } catch (e) {}
    return a;
  },
  deleteAlbum(req, me) {
    need(isAdmin(me));
    const a = getOne('albums', req.id); need(a, 'NOT_FOUND');
    need(!readAll('photos').some(p => p.albumId === a.id), 'NOT_EMPTY');
    if (a.folderId) { try { DriveApp.getFolderById(a.folderId).setTrashed(true); } catch (x) {} }
    del('albums', a.id);
    return true;
  },

  saveMember(req, me) {
    need(isAdmin(me));
    const x = req.member || {}, email = String(x.email || '').trim().toLowerCase();
    need(/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email), 'BAD_EMAIL');
    const m = getOne('members', email) || { id: email, email: email, addedAt: Date.now() };
    need(!m.owner || x.role === undefined || x.role === 'admin', 'OWNER_FIXED');
    if ('name' in x) m.name = String(x.name || '').slice(0, 30);
    if ('relation' in x) m.relation = String(x.relation || '').slice(0, 10);
    if ('role' in x && !m.owner) { need(ROLES.indexOf(x.role) >= 0, 'BAD_ROLE'); m.role = x.role; }
    if (!m.role) m.role = 'viewer';
    if ('canDownload' in x && !m.owner) m.canDownload = x.canDownload === true;
    if ('notify' in x) m.notify = x.notify !== false;
    put('members', m);
    return m;
  },
  removeMember(req, me) {
    need(isAdmin(me));
    const m = getOne('members', String(req.email || '').toLowerCase()); need(m, 'NOT_FOUND');
    need(!m.owner && OWNERS.map(s => s.toLowerCase()).indexOf(m.email) < 0, 'OWNER_FIXED');
    del('members', m.id);
    readAll('albums').forEach(a => { if ((a.allowed || []).indexOf(m.email) >= 0) { a.allowed = a.allowed.filter(e => e !== m.email); put('albums', a); } });
    return true;
  },

  setConfig(req, me) {
    need(isAdmin(me));
    const inc = req.cfg || {}, c = getOne('config', 'baby') || { id: 'baby' };
    if ('babyName' in inc) c.babyName = String(inc.babyName || '').slice(0, 20);
    if ('fullName' in inc) c.fullName = String(inc.fullName || '').slice(0, 30);
    if ('birthday' in inc) c.birthday = /^\d{4}-\d{2}-\d{2}$/.test(inc.birthday || '') ? inc.birthday : '';
    if ('gender' in inc) c.gender = ['girl', 'boy'].indexOf(inc.gender) >= 0 ? inc.gender : '';
    if ('avatarMode' in inc) c.avatarMode = inc.avatarMode === 'photo' ? 'photo' : 'growth';
    if ('avatar' in inc) c.avatar = (typeof inc.avatar === 'string' && inc.avatar.length < 45000 && /^data:image\/jpeg;base64,[A-Za-z0-9+\/=]+$/.test(inc.avatar)) ? inc.avatar : '';
    if ('showPresence' in inc) c.showPresence = inc.showPresence !== false;
    if ('setupDone' in inc) c.setupDone = !!inc.setupDone;
    put('config', c);
    return c;
  },

  log(req, me) { log(me, req.text || ''); return true; },
  activity(req, me) {
    need(isAdmin(me));
    const s = sheet('activity'), last = s.getLastRow(); if (!last) return [];
    const from = Math.max(1, last - 199);
    return s.getRange(from, 2, last - from + 1, 1).getValues().map(r => JSON.parse(r[0])).reverse();
  },

  deleteForever(req, me) {
    need(isAdmin(me));
    const p = getOne('photos', req.id); need(p, 'NOT_FOUND');
    trashFiles(p);   // 檔案先進你雲端硬碟的垃圾桶（Google 30 天後自動清空），多一層保險
    readAll('comments').filter(c => c.photoId === p.id).forEach(c => del('comments', c.id));
    del('photos', p.id);
    return true;
  },
};

function purgeOldTrash(me) {
  const old = readAll('photos').filter(p => p.state === 'trash' && p.trashedAt && Date.now() - p.trashedAt > 30 * 864e5);
  if (!old.length) return;
  const l = LockService.getScriptLock();
  if (!l.tryLock(5000)) return;
  try { old.forEach(p => { trashFiles(p); del('photos', p.id); }); log(me, '自動清除回收桶中超過 30 天的 ' + old.length + ' 個檔案'); }
  finally { l.releaseLock(); }
}
