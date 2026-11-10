/**
 * 會計學刷題站 雲端同步(Google 試算表 + Apps Script)
 * 使用方式見「雲端設定說明.md」。
 * 部署時請選:執行身分=我、存取權限=所有人。
 */
const SALT = '請改成你自己隨便打的一串字';   // ← 部署前改掉,之後不要再改(改了所有人密碼都會失效)
const SHEET = '學生紀錄';
const CHUNK = 40000, NCHUNK = 12, FIRST_CHUNK_COL = 11;   // 資料分段存在 K 欄之後
const MAX_FAILS = 5, LOCK_MIN = 10;

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet() { return out({ ok: true, msg: '會計學刷題站雲端運作中' }); }

function sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.getRange(1, 1, 1, 10).setValues([['名字', '密碼雜湊(勿改)', '輸錯次數', '鎖定到', '最後更新', '錯題數', '已精熟', '作答次數', '答對率%', '目前科目']]);
    sh.getRange(1, FIRST_CHUNK_COL).setValue('資料(勿改)');
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange(1, 1, 1, 10).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function hashPin(name, pin) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, name + ':' + pin + ':' + SALT, Utilities.Charset.UTF_8);
  return raw.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (_) { return out({ ok: false, err: 'busy' }); }
  try { return out(handle(JSON.parse(e.postData.contents))); }
  catch (err) { return out({ ok: false, err: 'server' }); }
  finally { lock.releaseLock(); }
}

function handle(b) {
  const name = String(b.name || '').normalize('NFKC').trim().slice(0, 20);
  const pin = String(b.pin || '');
  if (!name || !/^\d{4}$/.test(pin) || (b.a !== 'login' && b.a !== 'save')) return { ok: false, err: 'bad' };
  const sh = sheet(), last = sh.getLastRow(), now = Date.now(), h = hashPin(name, pin);
  const names = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
  const idx = names.indexOf(name), row = idx + 2;

  if (b.a === 'login') {
    if (idx < 0) {                                   // 新帳號
      const r = last + 1;
      sh.getRange(r, 1, 1, 2).setValues([[name, h]]);
      sh.getRange(r, 3, 1, 2).setValues([[0, 0]]);
      writeData(sh, r, b.data || { paper: 'B', wrong: {}, mastered: {}, log: [] });
      return { ok: true, created: true, data: null };
    }
    const rv = sh.getRange(row, 1, 1, 4).getValues()[0];
    if (Number(rv[3]) > now) return { ok: false, err: 'locked', wait: Math.ceil((Number(rv[3]) - now) / 60000) };
    if (rv[1] === '') {                              // 老師清除過密碼:第一次登入重新設定,紀錄保留
      sh.getRange(row, 2).setValue(h);
    } else if (rv[1] !== h) {
      let f = Number(rv[2]) + 1, lockTo = 0;
      if (f >= MAX_FAILS) { f = 0; lockTo = now + LOCK_MIN * 60000; }
      sh.getRange(row, 3, 1, 2).setValues([[f, lockTo]]);
      return { ok: false, err: 'pin' };
    }
    sh.getRange(row, 3, 1, 2).setValues([[0, 0]]);
    return { ok: true, data: readData(sh, row) };
  }

  // save
  if (idx < 0) return { ok: false, err: 'auth' };
  const cur = sh.getRange(row, 2).getValue();
  if (cur !== h) return { ok: false, err: 'auth' };
  writeData(sh, row, b.data || {});
  return { ok: true };
}

function readData(sh, row) {
  const cells = sh.getRange(row, FIRST_CHUNK_COL, 1, NCHUNK).getValues()[0];
  const s = cells.join('');
  if (!s) return null;
  try { return JSON.parse(s); } catch (_) { return null; }
}

function writeData(sh, row, d) {
  d = { paper: d.paper === 'C' ? 'C' : 'B', wrong: d.wrong || {}, mastered: d.mastered || {}, log: d.log || [] };
  let s = JSON.stringify(d);
  while (s.length > CHUNK * NCHUNK && d.log.length > 50) { d.log = d.log.slice(Math.floor(d.log.length * 0.2)); s = JSON.stringify(d); }
  const cells = [];
  for (let i = 0; i < NCHUNK; i++) cells.push(s.substr(i * CHUNK, CHUNK));
  const nl = d.log.length, ok = d.log.filter(function (l) { return l.ok; }).length;
  sh.getRange(row, 5, 1, 6).setValues([[new Date(), Object.keys(d.wrong).length, Object.keys(d.mastered).length, nl, nl ? Math.round(ok * 100 / nl) : 0, '會計' + d.paper]]);
  sh.getRange(row, FIRST_CHUNK_COL, 1, NCHUNK).setValues([cells]);
}

/* ---- 老師選單:在試算表上方出現「刷題站管理」 ---- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('刷題站管理')
    .addItem('讓某位學生重設密碼(保留紀錄)', 'resetPin')
    .addItem('刪除某位學生的帳號與紀錄', 'deleteUser')
    .addToUi();
}
function findRow_(name) {
  const sh = sheet(), last = sh.getLastRow();
  if (last < 2) return 0;
  const names = sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); });
  const i = names.indexOf(String(name).normalize('NFKC').trim());
  return i < 0 ? 0 : i + 2;
}
function resetPin() {
  const ui = SpreadsheetApp.getUi(), r = ui.prompt('要重設密碼的學生名字:');
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const row = findRow_(r.getResponseText());
  if (!row) return ui.alert('找不到這個名字');
  const sh = sheet();
  sh.getRange(row, 2).setValue('');
  sh.getRange(row, 3, 1, 2).setValues([[0, 0]]);
  ui.alert('已清除。請該學生用原本的名字、輸入「新密碼」登入一次,密碼就會改成新的,錯題紀錄都還在。\n(在他重設前,別人也可能搶先用這個名字登入,請告知學生馬上處理。)');
}
function deleteUser() {
  const ui = SpreadsheetApp.getUi(), r = ui.prompt('要刪除的學生名字(紀錄無法復原):');
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const row = findRow_(r.getResponseText());
  if (!row) return ui.alert('找不到這個名字');
  sheet().deleteRow(row);
  ui.alert('已刪除。');
}
