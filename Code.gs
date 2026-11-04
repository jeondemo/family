const SHEET_NAME = 'videos';
const HEAD = ['id', 'youtubeId', 'title', 'date', 'precision', 'who', 'albums', 'memo', 'createdAt'];
const PRECISIONS = ['day', 'month', 'year'];
const TZ = 'Asia/Seoul';

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange(2, 1, sh.getMaxRows() - 1, HEAD.length).setNumberFormat('@');
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TOKEN_SECRET')) props.setProperty('TOKEN_SECRET', Utilities.getUuid() + Utilities.getUuid());
  if (!props.getProperty('FAMILY_PIN')) props.setProperty('FAMILY_PIN', '0000');
}

function setPin() {
  const newPin = '1234';
  if (!/^\d{4}$/.test(newPin)) throw new Error('비밀번호는 숫자 4자리여야 합니다.');
  PropertiesService.getScriptProperties().setProperty('FAMILY_PIN', newPin);
}

function doGet() {
  return json({ ok: true, app: 'family-video' });
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    out = route(req);
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return json(out);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function route(req) {
  if (req.action === 'login') return login(String(req.pin || ''));
  if (!validToken(req.token)) return { ok: false, error: 'AUTH' };
  switch (req.action) {
    case 'list': return { ok: true, videos: listVideos() };
    case 'check': return checkVideo(String(req.url || ''));
    case 'add': return saveVideo(req.video, false);
    case 'update': return saveVideo(req.video, true);
    case 'remove': return removeVideo(String(req.id || ''));
  }
  return { ok: false, error: 'UNKNOWN_ACTION' };
}

function makeToken(pin) {
  const secret = PropertiesService.getScriptProperties().getProperty('TOKEN_SECRET');
  const sig = Utilities.computeHmacSha256Signature(pin, secret);
  return Utilities.base64EncodeWebSafe(sig);
}

function validToken(token) {
  if (!token) return false;
  const pin = PropertiesService.getScriptProperties().getProperty('FAMILY_PIN');
  return token === makeToken(pin);
}

function login(pin) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('fails') || 0);
  if (fails >= 10) return { ok: false, error: 'LOCKED' };
  const real = PropertiesService.getScriptProperties().getProperty('FAMILY_PIN');
  if (pin !== real) {
    cache.put('fails', String(fails + 1), 600);
    return { ok: false, error: 'WRONG_PIN' };
  }
  cache.remove('fails');
  return { ok: true, token: makeToken(real) };
}

function sheet() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('setup을 먼저 실행해 주세요.');
  return sh;
}

function cellText(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v === null || v === undefined ? '' : v);
}

function listVideos() {
  const sh = sheet();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, HEAD.length).getValues();
  return rows.filter(r => r[0]).map(r => {
    const o = {};
    HEAD.forEach((h, i) => { o[h] = cellText(r[i]); });
    o.albums = o.albums ? o.albums.split(',').map(s => s.trim()).filter(Boolean) : [];
    return o;
  });
}

function parseYoutubeId(url) {
  const s = String(url).trim();
  if (/^[\w-]{11}$/.test(s)) return s;
  const m = s.match(/(?:youtu\.be\/|[?&]v=|\/shorts\/|\/embed\/|\/live\/)([\w-]{11})/);
  return m ? m[1] : '';
}

function checkVideo(url) {
  const id = parseYoutubeId(url);
  if (!id) return { ok: false, error: 'BAD_URL' };
  const res = UrlFetchApp.fetch('https://www.youtube.com/oembed?format=json&url=' + encodeURIComponent('https://www.youtube.com/watch?v=' + id), { muteHttpExceptions: true });
  const code = res.getResponseCode();
  if (code === 200) {
    const d = JSON.parse(res.getContentText());
    return { ok: true, youtubeId: id, title: d.title || '' };
  }
  if (code === 401 || code === 403) return { ok: false, error: 'PRIVATE', youtubeId: id };
  return { ok: false, error: 'NOT_FOUND', youtubeId: id };
}

function safe(v, max) {
  let s = String(v === null || v === undefined ? '' : v).trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function cleanVideo(v) {
  if (!v) throw new Error('영상 정보가 없습니다.');
  const youtubeId = parseYoutubeId(v.youtubeId || '');
  if (!youtubeId) throw new Error('유튜브 링크를 확인해 주세요.');
  const title = safe(v.title, 100);
  if (!title) throw new Error('제목을 입력해 주세요.');
  const precision = PRECISIONS.indexOf(v.precision) >= 0 ? v.precision : 'day';
  const date = String(v.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('촬영일을 확인해 주세요.');
  const albums = (Array.isArray(v.albums) ? v.albums : []).map(a => safe(a, 20).replace(/,/g, ' ')).filter(Boolean).join(',');
  return {
    youtubeId: youtubeId,
    title: title,
    date: date,
    precision: precision,
    who: safe(v.who, 10),
    albums: albums,
    memo: safe(v.memo, 500)
  };
}

function findRow(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return -1;
}

function saveVideo(v, isUpdate) {
  const c = cleanVideo(v);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet();
    if (isUpdate) {
      const row = findRow(sh, String(v.id || ''));
      if (row < 0) return { ok: false, error: 'NOT_FOUND' };
      const created = sh.getRange(row, 9).getValue();
      sh.getRange(row, 1, 1, HEAD.length).setNumberFormat('@').setValues([[v.id, c.youtubeId, c.title, c.date, c.precision, c.who, c.albums, c.memo, cellText(created)]]);
      return { ok: true, id: v.id };
    }
    const id = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    const now = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss");
    const row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, HEAD.length).setNumberFormat('@').setValues([[id, c.youtubeId, c.title, c.date, c.precision, c.who, c.albums, c.memo, now]]);
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

function removeVideo(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet();
    const row = findRow(sh, id);
    if (row < 0) return { ok: false, error: 'NOT_FOUND' };
    sh.deleteRow(row);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}
