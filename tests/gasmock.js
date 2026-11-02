// 模擬 Google Apps Script 服務，用實際資料範本測試 Code.gs
const fs = require('fs');
const vm = require('vm');
const crypto = require('crypto');

function makeEnv(workbookPath, opts = {}) {
  const wb = JSON.parse(fs.readFileSync(workbookPath, 'utf8'));
  const log = { mails: [], files: [], created: [], fetches: [], moved: [], trashed: [] };
  const toCell = v => (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v)) ? new Date(v + 'T00:00:00+08:00') : v; // 模擬試算表把日期字串轉成日期

  class Range {
    constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
    getValues() { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) { const v = (this.sh.data[this.r - 1 + i] || [])[this.c - 1 + j]; row.push(v === undefined ? '' : v); } out.push(row); } return out; }
    setValues(vals) { if (vals.length !== this.nr || vals[0].length !== this.nc) throw new Error(`setValues 尺寸不符 ${vals.length}x${vals[0].length} vs ${this.nr}x${this.nc}`); vals.forEach((row, i) => row.forEach((v, j) => this.sh.set(this.r + i, this.c + j, toCell(v)))); return this; }
    setValue(v) { this.sh.set(this.r, this.c, toCell(v)); return this; }
  }
  // 格式設定類方法（merge、setFontWeight…）在模擬環境不做事，只回傳自己以便串接
  const chain = obj => { const p = new Proxy(obj, { get: (t, k) => (k in t) ? (typeof t[k] === 'function' ? (...a) => { const r = t[k](...a); return r === t ? p : r; } : t[k]) : (() => p) }); return p; };
  class Sheet {
    constructor(name, data) { this.name = name; this.data = data.map(r => r.slice()); }
    getName() { return this.name; }
    set(r, c, v) { while (this.data.length < r) this.data.push([]); const row = this.data[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; }
    getLastRow() { let n = this.data.length; while (n && this.data[n - 1].every(v => v === '' || v === undefined)) n--; return n; }
    getLastColumn() { return this.data.reduce((m, r) => Math.max(m, r.length), 0); }
    getDataRange() { return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
    getRange(r, c, nr = 1, nc = 1) { if (r < 1 || c < 1 || nr < 1 || nc < 1) throw new Error(`getRange 參數錯誤 ${r},${c},${nr},${nc}`); return chain(new Range(this, r, c, nr, nc)); }
    setName(n) { this.name = n; return this; }
    setColumnWidth() { return this; }
    setFrozenRows() { return this; }
  }
  const sheets = {};
  Object.keys(wb).forEach(n => sheets[n] = new Sheet(n, wb[n]));
  const ss = { getName: () => '測試資料庫', getSheets: () => Object.values(sheets), getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new Sheet(n, [])) };

  function fmtDate(d, tz, p) {
    const parts = {};
    new Intl.DateTimeFormat('en-GB', { timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false, weekday: 'short' })
      .formatToParts(d).forEach(x => parts[x.type] = x.value);
    const u = { Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6, Sun: 7 }[parts.weekday];
    return p.replace('yyyy', parts.year).replace('MM', parts.month).replace('dd', parts.day).replace('HH', parts.hour === '24' ? '00' : parts.hour).replace('mm', parts.minute).replace('ss', parts.second).replace('u', u);
  }
  function folder(name) {
    const kids = {};
    return {
      name,
      getFoldersByName: n => { const f = kids[n]; let used = false; return { hasNext: () => !!f && !used, next: () => { used = true; return f; } }; },
      createFolder: n => (kids[n] = folder(n)),
      createFile: blob => { log.files.push({ folder: name, name: blob.name, size: blob.bytes.length }); return { getUrl: () => 'https://drive.google.com/file/d/' + crypto.randomBytes(6).toString('hex') + '/view' }; }
    };
  }
  const root = folder('巡檢照片');
  const cache = new Map();
  const ctx = {
    console,
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, openById: () => ss, flush: () => {},
      create: name => { const sh = new Sheet('工作表1', []); const id = 'tmp' + log.created.length; const t = { getId: () => id, getName: () => name, getSheets: () => [sh] }; log.created.push({ id, name, sheet: sh }); return t; } },
    UrlFetchApp: { fetch: (url, o) => { log.fetches.push(url); return { getResponseCode: () => 200, getBlob: () => ({ getBytes: () => Array.from(Buffer.from('PK-fake-xlsx')) }) }; } },
    ScriptApp: { getOAuthToken: () => 'token' },
    Utilities: {
      formatDate: fmtDate, getUuid: () => crypto.randomUUID(), sleep: () => {},
      base64Decode: s => Array.from(Buffer.from(s, 'base64')), base64Encode: b => Buffer.from(b).toString('base64'), newBlob: (bytes, mime, name) => ({ bytes, mime, name })
    },
    CacheService: { getScriptCache: () => ({ get: k => cache.has(k) ? cache.get(k) : null, put: (k, v) => cache.set(k, v), remove: k => cache.delete(k) }) },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (opts.props || { PHOTO_FOLDER_ID: 'root' })[k] || null }) },
    DriveApp: { getFolderById: () => root, getRootFolder: () => root, getFileById: id => ({ moveTo: () => log.moved.push(id), setTrashed: v => log.trashed.push(id) }) },
    MailApp: { getRemainingDailyQuota: () => 100, sendEmail: m => log.mails.push(m) },
    HtmlService: { createHtmlOutputFromFile: () => ({ setTitle() { return this; }, addMetaTag() { return this; } }) }
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(__dirname + '/../src/Code.gs', 'utf8'), ctx, { filename: 'Code.gs' });
  return { ctx, sheets, log };
}
module.exports = { makeEnv };
