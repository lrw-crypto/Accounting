/**
 * Install this code in a Google spreadsheet's bound Apps Script project.
 * Run setup() once from that project before deploying the web app.
 */
const SITE_TABS = Object.freeze({
  economy: '經濟學',
  Digital: '數位科技概論',
  Accounting: '會計學',
  Business: '商業概論',
  testmath: '數學',
});
const LEGACY_MATH_SALT = '請改成你自己隨便打的一串字';
const CHUNK_SIZE = 35000;
const CHUNKS = 10;
const HEADERS = ['學生代稱', '最近登入', '最近作答', '待複習錯題', '已克服', '近 6,000 筆作答數', '近 6,000 筆答對率', '目前卷別', '登入驗證', '連續輸錯', '鎖定至'];

function setup() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!sheet) throw new Error('請在目標試算表的擴充功能選單開啟 Apps Script。');
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SHEET_ID') && props.getProperty('SHEET_ID') !== sheet.getId()) {
    throw new Error('這個專案已連結另一份試算表。');
  }
  props.setProperty('SHEET_ID', sheet.getId());
  if (!props.getProperty('SIGNING_KEY')) {
    props.setProperty('SIGNING_KEY', Utilities.getUuid() + Utilities.getUuid());
  }
  Object.keys(SITE_TABS).forEach(site => openSiteTab_(sheet, site));
  migrateMath_(sheet);
  SpreadsheetApp.flush();
}

function openSiteTab_(book, site) {
  const name = SITE_TABS[site];
  const existing = book.getSheetByName(name);
  if (existing) {
    if (existing.getRange(1, 9).getValue() !== HEADERS[8]) {
      throw new Error('同名分頁已存在，請先將原分頁改名。');
    }
    return existing;
  }
  const tab = book.insertSheet(name);
  const headers = HEADERS.concat(Array.from({length: CHUNKS}, (_, i) => '同步資料 ' + (i + 1)));
  tab.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#2457a7').setWrap(true);
  tab.setFrozenRows(1);
  tab.setRowHeight(1, 38);
  tab.setColumnWidth(1, 145);
  tab.setColumnWidths(2, 2, 156);
  tab.setColumnWidths(4, 5, 108);
  tab.getRange('A:A').setNumberFormat('@');
  tab.getRange('B:C').setNumberFormat('yyyy/mm/dd hh:mm');
  tab.getRange('G:G').setNumberFormat('0%');
  tab.hideColumns(9, headers.length - 8);
  return tab;
}

function response_(body) {
  return ContentService.createTextOutput(JSON.stringify(body)).setMimeType(ContentService.MimeType.JSON);
}

function hasSite_(site) {
  return Object.prototype.hasOwnProperty.call(SITE_TABS, site);
}

function doGet(e) {
  const site = e?.parameter?.site || 'testmath';
  const configured = Boolean(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  return response_({ ok: configured && hasSite_(site), service: 'quiz-sync-v1', site: hasSite_(site) ? site : null });
}

function doPost(e) {
  const site = e?.parameter?.site || 'testmath';
  const input = e && e.postData && e.postData.contents;
  if (!hasSite_(site) || typeof input !== 'string' || input.length > 550000) return response_({ok: false, err: 'bad'});
  let request;
  try { request = JSON.parse(input); } catch (_) { return response_({ok: false, err: 'bad'}); }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(12000)) return response_({ok: false, err: 'busy'});
  try {
    return response_(handle_(site, request));
  } catch (_) {
    return response_({ok: false, err: 'server'});
  } finally {
    lock.releaseLock();
  }
}

function handle_(site, request) {
  if (!request || typeof request.name !== 'string' || typeof request.pin !== 'string') return {ok: false, err: 'bad'};
  const name = request.name.normalize('NFKC').trim().toLowerCase();
  const pin = request.pin;
  if (!name || name.length > 20 || /[\u0000-\u001f\u007f]/.test(name) || !/^\d{4}$/.test(pin)
      || !['login', 'save'].includes(request.a)) return {ok: false, err: 'bad'};

  const props = PropertiesService.getScriptProperties();
  const bookId = props.getProperty('SHEET_ID');
  const signingKey = props.getProperty('SIGNING_KEY');
  if (!bookId || !signingKey) return {ok: false, err: 'setup'};
  const signature = Utilities.computeHmacSha256Signature(JSON.stringify([site, name, pin]), signingKey)
    .map(byte => ('0' + (byte & 255).toString(16)).slice(-2)).join('');
  const tab = openSiteTab_(SpreadsheetApp.openById(bookId), site);
  const last = tab.getLastRow();
  const displayNames = last > 1 ? tab.getRange(2, 1, last - 1, 1).getDisplayValues() : [];
  const index = displayNames.findIndex(row => row[0] === name);
  const now = Date.now();
  const row = index < 0 ? last + 1 : index + 2;
  let previous = [];
  if (index >= 0) {
    previous = tab.getRange(row, 1, 1, HEADERS.length + CHUNKS).getValues()[0];
    if (Number(previous[10]) > now) return {ok: false, err: 'locked', wait: Math.ceil((Number(previous[10]) - now) / 60000)};
    if (previous[8] !== signature && !matchesLegacyMath_(site, name, pin, previous[8])) {
      const failures = Number(previous[9] || 0) + 1;
      tab.getRange(row, 10, 1, 2).setValues([[failures >= 5 ? 0 : failures, failures >= 5 ? now + 600000 : 0]]);
      SpreadsheetApp.flush();
      return {ok: false, err: 'pin'};
    }
  } else if (request.a !== 'login') {
    return {ok: false, err: 'auth'};
  }

  const encoded = previous.slice(HEADERS.length).join('');
  const stored = encoded ? JSON.parse(encoded) : {};
  const merged = mergeProgress_(stored, request.data, now);
  const recent = merged.log;
  const correct = recent.reduce((count, attempt) => count + attempt.ok, 0);
  const data = JSON.stringify(merged);
  if (data.length > CHUNKS * CHUNK_SIZE) return {ok: false, err: 'size'};
  const chunks = Array.from({length: CHUNKS}, (_, i) => data.slice(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE));
  if (row > tab.getMaxRows()) tab.insertRowsAfter(tab.getMaxRows(), 100);
  const values = [
    "'" + name,
    request.a === 'login' ? new Date(now) : previous[1] || new Date(now),
    recent.length ? new Date(recent[recent.length - 1].t) : '',
    Object.keys(merged.wrong).length, Object.keys(merged.mastered).length,
    recent.length, recent.length ? correct / recent.length : 0, merged.paper || '',
    signature, 0, 0,
  ];
  tab.getRange(row, 1, 1, values.length + chunks.length).setValues([values.concat(chunks)]);
  SpreadsheetApp.flush();
  return {ok: true, created: index < 0, data: merged};
}

function matchesLegacyMath_(site, name, pin, stored) {
  if (site !== 'testmath' || typeof stored !== 'string' || !stored.startsWith('legacy:')) return false;
  if (stored === 'legacy:reset') return true;
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, name + ':' + pin + ':' + LEGACY_MATH_SALT, Utilities.Charset.UTF_8)
    .map(byte => ('0' + (byte & 255).toString(16)).slice(-2)).join('');
  return stored === 'legacy:' + digest;
}

function migrateMath_(book) {
  const legacy = book.getSheetByName('學生紀錄');
  if (!legacy || legacy.getLastRow() < 2) return;
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('LEGACY_MATH_IMPORTED') === 'yes') return;
  const math = openSiteTab_(book, 'testmath');
  const count = legacy.getLastRow() - 1;
  const old = legacy.getRange(2, 1, count, 22).getValues();
  const existing = math.getLastRow() > 1 ? math.getRange(2, 1, math.getLastRow() - 1, 1).getDisplayValues().flat() : [];
  const seen = new Set(existing);
  const values = [];
  for (const row of old) {
    const name = String(row[0] || '').normalize('NFKC').trim().toLowerCase();
    if (!name || seen.has(name)) continue;
    const raw = row.slice(10).join('');
    const progress = mergeProgress_({}, raw ? JSON.parse(raw) : {}, Date.now());
    const data = JSON.stringify(progress);
    if (data.length > CHUNKS * CHUNK_SIZE) throw new Error('舊紀錄太大，請檢查數學資料。');
    const count = progress.log.length;
    const correct = progress.log.reduce((n, attempt) => n + attempt.ok, 0);
    const recent = count ? new Date(progress.log[count - 1].t) : row[4] || '';
    const digest = row[1] ? 'legacy:' + row[1] : 'legacy:reset';
    const cells = Array.from({length: CHUNKS}, (_, i) => data.slice(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE));
    values.push(["'" + name, row[4] || '', recent, Object.keys(progress.wrong).length,
      Object.keys(progress.mastered).length, count, count ? correct / count : 0, progress.paper || '',
      digest, row[2] || 0, row[3] || 0, ...cells]);
    seen.add(name);
  }
  if (values.length) {
    const start = math.getLastRow() + 1;
    const extra = start + values.length - 1 - math.getMaxRows();
    if (extra > 0) math.insertRowsAfter(math.getMaxRows(), extra);
    math.getRange(start, 1, values.length, HEADERS.length + CHUNKS).setValues(values);
    SpreadsheetApp.flush();
  }
  props.setProperty('LEGACY_MATH_IMPORTED', 'yes');
}

function mergeProgress_(stored, incoming, now) {
  const target = {paper: undefined, wrong: Object.create(null), mastered: Object.create(null), log: []};
  const latest = new Map();
  const attempts = new Map();
  const validId = id => typeof id === 'string' && id.length >= 1 && id.length <= 65 && /^[\w.:-]+$/.test(id)
    && !['__proto__', 'constructor', 'prototype'].includes(id);
  const validTime = t => Number.isSafeInteger(t) && t > 0 && t <= now + 600000;
  for (const payload of [stored, incoming]) {
    if (!payload || typeof payload !== 'object') continue;
    if (typeof payload.paper === 'string' && /^[ABC]$/.test(payload.paper)) target.paper = payload.paper;
    for (const state of ['wrong', 'mastered']) {
      const entries = payload[state];
      if (!entries || typeof entries !== 'object' || Array.isArray(entries)) continue;
      for (const [id, item] of Object.entries(entries).slice(0, 4000)) {
        if (!validId(id) || !item || !validTime(item.t) || !Number.isSafeInteger(item.n) || item.n < 1) continue;
        const old = latest.get(id);
        if (!old || item.t > old.t || (item.t === old.t && state === 'mastered')) {
          latest.set(id, {state, t: item.t, n: item.n});
        }
      }
    }
    if (!Array.isArray(payload.log)) continue;
    for (const item of payload.log.slice(-6000)) {
      if (!item || !validId(item.id) || !validTime(item.t) || ![0, 1, true, false].includes(item.ok)) continue;
      const mode = ['practice', 'wrong', 'exam', 'review'].includes(item.m) ? item.m : 'practice';
      const key = item.id + '@' + item.t;
      if (!attempts.has(key)) attempts.set(key, {id: item.id, t: item.t, ok: Number(item.ok), m: mode});
    }
  }
  for (const [id, item] of latest) target[item.state][id] = {t: item.t, n: item.n};
  target.log = [...attempts.values()].sort((a, b) => a.t - b.t).slice(-6000);
  while (JSON.stringify(target).length > CHUNKS * CHUNK_SIZE && target.log.length > 0) {
    target.log.splice(0, Math.max(1, Math.ceil(target.log.length / 5)));
  }
  return target;
}
