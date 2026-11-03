// 模擬 Google Apps Script 服務，用實際資料範本測試 Code.gs
const fs = require('fs');
const vm = require('vm');
const crypto = require('crypto');
const SRC = __dirname + '/../src';

function makeEnv(workbookPath, opts = {}) {
  const wb = JSON.parse(fs.readFileSync(workbookPath, 'utf8'));
  const log = { mails: [], files: [], created: [], fetches: [], moved: [], trashed: [], images: [] };
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
    setRowHeight() { return this; }
    insertImage(blob, c, r) { log.images.push({ sheet: this.name, c, r, size: (blob.bytes || []).length }); let w = 600, h = 450; const img = { getWidth: () => w, getHeight: () => h, setWidth: v => { w = v; return img; }, setHeight: v => { h = v; return img; } }; return img; }
    setFrozenRows() { return this; }
  }
  const sheets = {};
  Object.keys(wb).forEach(n => sheets[n] = new Sheet(n, wb[n]));
  const ss = { getId: () => 'ss1', getName: () => '測試資料庫', getSheets: () => Object.values(sheets), getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new Sheet(n, [])) };

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
      getId: () => 'folder-' + name,
      createFolder: n => (kids[n] = folder(n)),
      createFile: blob => { log.files.push({ folder: name, name: blob.name, size: blob.bytes.length }); return { getUrl: () => 'https://drive.google.com/file/d/' + crypto.randomBytes(6).toString('hex') + '/view?usp=drivesdk' }; }
    };
  }
  const root = folder('巡檢照片');
  const cache = new Map();
  const props = Object.assign({}, opts.props || { PHOTO_FOLDER_ID: 'root' });
  const ctx = {
    console,
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, openById: () => ss, flush: () => {},
      create: name => { const sh = new Sheet('工作表1', []); const id = 'tmp' + log.created.length; const list = [sh]; const t = { getId: () => id, getName: () => name, getSheets: () => list, insertSheet: n => { const x = new Sheet(n, []); list.push(x); return x; } }; log.created.push({ id, name, sheet: sh, sheets: list }); return t; } },
    UrlFetchApp: { fetch: (url, o) => { log.fetches.push(url); const isThumb = url.indexOf('thumbnail') >= 0; return { getResponseCode: () => 200, getHeaders: () => ({ 'Content-Type': isThumb ? 'image/jpeg' : 'application/vnd.openxmlformats' }), getBlob: () => ({ bytes: [1, 2, 3], getBytes: () => Array.from(Buffer.from('PK-fake-xlsx')) }) }; } },
    ScriptApp: { getOAuthToken: () => 'token' },
    Utilities: {
      Charset: { UTF_8: 'UTF-8' }, formatDate: fmtDate, getUuid: () => crypto.randomUUID(), sleep: () => {},
      base64Decode: s => Array.from(Buffer.from(s, 'base64')), base64Encode: b => Buffer.from(typeof b === 'string' ? Buffer.from(b, 'utf8') : b).toString('base64'), newBlob: (bytes, mime, name) => ({ bytes, mime, name })
    },
    CacheService: { getScriptCache: () => ({ get: k => cache.has(k) ? cache.get(k) : null, put: (k, v) => cache.set(k, v), remove: k => cache.delete(k) }) },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
    DriveApp: { getFolderById: () => root, getRootFolder: () => root, getFileById: id => ({ moveTo: () => log.moved.push(id), setTrashed: v => log.trashed.push(id), getParents: () => { let u = false; return { hasNext: () => !u, next: () => { u = true; return root; } }; } }) },
    MailApp: { getRemainingDailyQuota: () => 100, sendEmail: m => log.mails.push(m) },
    HtmlService: {
      createHtmlOutputFromFile: () => ({ setTitle() { return this; }, addMetaTag() { return this; } }),
      createTemplateFromFile: name => {
        const raw = fs.readFileSync(SRC + '/' + name + '.html', 'utf8');
        const t = { getRawContent: () => raw, evaluate: () => {
          const html = raw.split('<?= appB64 ?>').join(String(t.appB64 || ''));
          return { getContent: () => html, setTitle() { return this; }, addMetaTag() { return this; } }; } };
        return t; } }
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(SRC + '/Code.gs', 'utf8'), ctx, { filename: 'Code.gs' });
  return { ctx, sheets, log, props };
}
module.exports = { makeEnv };
