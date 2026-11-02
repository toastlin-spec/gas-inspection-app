/**
 * 加油站巡檢 App — 後端（第一階段）
 *
 * 功能：密碼登入、首頁狀態、自主檢查送出（含照片）、SL-H015 查核與匯出 Excel、異常清單（唯讀）、管理設定、資料檢查。
 * 資料全部存在綁定的 Google 試算表；照片存在指令碼屬性 PHOTO_FOLDER_ID 指定的資料夾。
 *
 * 指令碼屬性（專案設定 → 指令碼屬性）：
 *   PHOTO_FOLDER_ID   「巡檢照片」資料夾 ID（必填）
 *   EXPORT_FOLDER_ID  「匯出暫存」資料夾 ID（第三階段匯出 Excel 使用）
 *   APP_NAME          App 名稱，預設「加油站巡檢」
 *   SHEET_ID          建議填寫：「加油站巡檢系統資料庫」網址 /d/ 後面那一段
 */

const TZ = 'Asia/Taipei';
const SESSION_SECONDS = 6 * 60 * 60; // CacheService 上限 6 小時
const REGIONS = ['北區', '中區', '南區'];
const RESULT_LABEL = { ok: '正常', ng: '異常', na: '不適用' };
const MAX_PHOTO_CHARS = 4 * 1024 * 1024;
const AUDIT_FORM = 'SL-H015查核表';
const AUDIT_LABEL = { ok: '符合', ng: '不符合', na: '不適用' };

const REQUIRED_HEADERS = {
  '站點': ['站點代碼', '站名', '區域', '組別', '站長', '站長Email', '啟用', '備註'],
  '人員': ['姓名', '角色', '負責區域', 'Email', '啟用', '備註'],
  '檢查項目': ['項目代碼', '表單類型', '版本', '頻率', '類別', '檢查內容', '判定標準', '法規依據', '風險等級', '配分', '排序', '啟用'],
  '檢查紀錄': ['紀錄編號', '表單類型', '版本', '站點代碼', '頻率', '檢查日期', '填報人', '送出時間', '異常數', '得分', '站長確認時間', '站長確認人'],
  '檢查明細': ['紀錄編號', '項目代碼', '結果', '得分', '說明', '照片連結'],
  '異常追蹤': ['異常編號', '紀錄編號', '站點代碼', '來源', '項目代碼', '異常說明', '異常照片', '風險等級', '狀態', '建立日期', '改善期限', '處置說明', '改善後照片', '結案人', '結案日期', '處置紀錄'],
  '設定': ['參數', '值', '單位', '說明']
};

/* ============================== 進入點 ============================== */

function doGet() {
  // 畫面程式放在 App.html，以原始內容（getRawContent）讀取並轉 Base64 傳給 Index.html，
  // 避免 Google 輸出網頁時改動程式內容。
  const page = HtmlService.createTemplateFromFile('Index');
  page.appB64 = Utilities.base64Encode(HtmlService.createTemplateFromFile('App').getRawContent(), Utilities.Charset.UTF_8);
  return page.evaluate()
    .setTitle(getProp_('APP_NAME') || '加油站巡檢')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** 開啟試算表時加上選單，方便管理員檢查資料結構。 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('巡檢系統')
    .addItem('補上缺少的欄位', 'menuSetupSheets')
    .addItem('檢查資料', 'menuCheckData')
    .addToUi();
}

function menuSetupSheets() {
  const added = setupSheets_();
  SpreadsheetApp.getUi().alert(added.length ? '已補上欄位：\n' + added.join('\n') : '欄位完整，不需要補上。');
}

function menuCheckData() {
  const list = checkData_();
  SpreadsheetApp.getUi().alert(list.length
    ? list.map(p => `【${p.level}】${p.sheet}：${p.msg}`).join('\n')
    : '沒有發現問題。');
}

/* ============================== 登入與工作階段 ============================== */

/** 第一步：確認密碼，回傳角色與第二步需要的選項。 */
function checkPassword(pw) {
  const role = roleForPassword_(pw);
  if (!role) {
    Utilities.sleep(1500); // 放慢連續猜密碼
    return { ok: false, message: '密碼不正確，請再試一次' };
  }
  const out = { ok: true, role: role };
  if (role === '站長') out.stations = activeStations_().map(publicStation_);
  if (role === '安保專員') {
    const all = activeStations_();
    out.officers = officers_().filter(o => o.on).map(o => ({
      name: o.name, regions: o.regions,
      count: all.filter(s => o.regions.indexOf(s.region) >= 0).length
    }));
  }
  return out;
}

/** 第二步：建立工作階段，回傳 token。 */
function login(req) {
  req = req || {};
  const role = roleForPassword_(req.pw);
  if (!role) { Utilities.sleep(1500); throw new Error('密碼不正確，請重新登入'); }
  const sess = { role: role, name: '', station: '', regions: [] };
  if (role === '站長') {
    const st = activeStations_().find(s => s.code === String(req.station || ''));
    if (!st) throw new Error('請選擇站點');
    const name = clean_(req.name, 30);
    if (!name) throw new Error('請輸入填報人姓名');
    sess.name = name; sess.station = st.code; sess.regions = [st.region];
  } else if (role === '安保專員') {
    const o = officers_().find(x => x.on && x.name === String(req.officer || ''));
    if (!o) throw new Error('請選擇你的姓名');
    sess.name = o.name; sess.regions = o.regions;
  } else {
    sess.name = '系統管理員'; sess.regions = REGIONS.slice();
  }
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('S_' + token, JSON.stringify(sess), SESSION_SECONDS);
  return { token: token, session: sess };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('S_' + token);
  return true;
}

function roleForPassword_(pw) {
  pw = String(pw === undefined || pw === null ? '' : pw).trim();
  if (!pw) return null;
  const s = settings_();
  const map = { '登入密碼_站長': '站長', '登入密碼_安保專員': '安保專員', '登入密碼_系統管理員': '系統管理員' };
  for (const key in map) {
    if (s[key] !== undefined && String(s[key]).trim() !== '' && String(s[key]).trim() === pw) return map[key];
  }
  return null;
}

function session_(token, roles) {
  const raw = token ? CacheService.getScriptCache().get('S_' + token) : null;
  if (!raw) throw new Error('SESSION_EXPIRED');
  const sess = JSON.parse(raw);
  if (roles && roles.indexOf(sess.role) < 0) throw new Error('沒有權限執行這個動作');
  return sess;
}

/* ============================== 首頁 ============================== */

function getHome(token) {
  const sess = session_(token);
  const stations = scopeStations_(sess);
  const codes = {};
  stations.forEach(s => { codes[s.code] = true; });
  const today = today_();
  const weekStart = weekStart_();
  const month = today.slice(0, 7);

  const status = {};
  stations.forEach(s => { status[s.code] = { daily: false, weekly: false, monthly: false }; });
  readTable_('檢查紀錄').rows.forEach(r => {
    const code = String(r['站點代碼']);
    if (!codes[code] || String(r['表單類型']) !== '自主檢查') return;
    const d = dateStr_(r['檢查日期']);
    const f = String(r['頻率']);
    if (f === '每日' && d === today) status[code].daily = true;
    if (f === '每週' && d >= weekStart) status[code].weekly = true;
    if (f === '每月' && d.slice(0, 7) === month) status[code].monthly = true;
  });

  const issues = issuesInScope_(codes);
  const counts = {
    open: issues.filter(i => i.status !== '已結案').length,
    overdue: issues.filter(i => i.overdue).length,
    verify: issues.filter(i => i.status === '待驗證').length
  };

  const items = items_().filter(i => i.on && i.form === '自主檢查').map(publicItem_);
  const audits = auditSummary_(codes);
  return {
    session: sess,
    today: today,
    stations: stations.map(s => Object.assign(publicStation_(s), status[s.code], { lastAudit: audits.last[s.code] || null, auditedThisQuarter: !!audits.quarter[s.code] })),
    counts: counts,
    audit: { quarter: quarterLabel_(today), done: Object.keys(audits.quarter).length, total: stations.length, recent: audits.recent.slice(0, 5) },
    items: items,
    riskDays: riskDays_()
  };
}

/* ============================== 自主檢查 ============================== */

/**
 * payload = { freq: '每日'|'每週'|'每月', answers: [{ code, result: 'ok'|'ng'|'na', note, photo }] }
 * photo 為前端壓縮後的 data URL（image/jpeg）。
 */
function submitCheck(token, payload) {
  const sess = session_(token, ['站長']);
  payload = payload || {};
  const freq = String(payload.freq || '');
  if (['每日', '每週', '每月'].indexOf(freq) < 0) throw new Error('檢查頻率不正確');

  const items = items_().filter(i => i.on && i.form === '自主檢查' && i.freq === freq);
  if (!items.length) throw new Error('這個頻率目前沒有啟用的檢查項目');
  const ans = {};
  (payload.answers || []).forEach(a => { if (a && a.code) ans[a.code] = a; });

  const missing = items.filter(i => !ans[i.code] || !RESULT_LABEL[ans[i.code].result]);
  if (missing.length) throw new Error('還有 ' + missing.length + ' 項未填寫');
  const noNote = items.filter(i => ans[i.code].result === 'ng' && !clean_(ans[i.code].note, 500));
  if (noNote.length) throw new Error('異常項目需要填寫說明：' + noNote.map(i => i.code).join('、'));

  const station = activeStations_().find(s => s.code === sess.station);
  if (!station) throw new Error('站點已停用，請聯絡系統管理員');

  const now = new Date();
  const today = today_();
  const stamp = fmt_(now, 'yyyy-MM-dd HH:mm:ss');
  const recordId = 'CK-' + fmt_(now, 'yyyyMMdd-HHmmss') + '-' + station.code;
  const days = riskDays_();
  const reporter = sess.name;

  // 照片先存（不需要鎖），失敗時整筆不寫入
  const photos = {};
  items.forEach(i => {
    const a = ans[i.code];
    if (a.result === 'ng' && a.photo) photos[i.code] = savePhoto_(a.photo, station.code, recordId + '_' + i.code);
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let issueRows = [];
  try {
    let seq = nextIssueSeq_(now);
    const results = [];
    items.forEach(i => {
      const a = ans[i.code];
      const note = clean_(a.note, 500);
      results.push({
        '紀錄編號': recordId, '項目代碼': i.code, '結果': RESULT_LABEL[a.result],
        '得分': '', '說明': note, '照片連結': photos[i.code] || ''
      });
      if (a.result === 'ng') {
        const id = 'ISS-' + fmt_(now, 'yyyyMM') + '-' + pad_(seq++, 3);
        issueRows.push({
          '異常編號': id, '紀錄編號': recordId, '站點代碼': station.code, '來源': '自主檢查',
          '項目代碼': i.code, '異常說明': note, '異常照片': photos[i.code] || '',
          '風險等級': i.risk, '狀態': '待處理', '建立日期': today,
          '改善期限': addDays_(today, days[i.risk] || 7), '處置說明': '', '改善後照片': '',
          '結案人': '', '結案日期': '',
          '處置紀錄': JSON.stringify([{ t: stamp, who: reporter + '（站長）', text: freq + '檢查通報異常' + (photos[i.code] ? '，附照片 1 張' : '') }])
        });
      }
    });
    appendObjects_('檢查紀錄', [{
      '紀錄編號': recordId, '表單類型': '自主檢查', '版本': 'v1', '站點代碼': station.code, '頻率': freq,
      '檢查日期': today, '填報人': reporter, '送出時間': stamp, '異常數': issueRows.length, '得分': '', '站長確認時間': ''
    }]);
    appendObjects_('檢查明細', results);
    appendObjects_('異常追蹤', issueRows);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  let mailed = false;
  try { mailed = notifyHighRisk_(station, issueRows, reporter, '自主檢查'); } catch (e) { console.warn('通知寄送失敗：' + e); }

  return {
    recordId: recordId, freq: freq, total: items.length, abnormal: issueRows.length,
    high: issueRows.filter(r => r['風險等級'] === '高').length, mailed: mailed
  };
}

function savePhoto_(dataUrl, stationCode, baseName) {
  const text = String(dataUrl || '');
  if (text.length > MAX_PHOTO_CHARS) throw new Error('照片太大，請重新拍攝');
  const m = /^data:(image\/(?:jpeg|png|webp));base64,(.+)$/.exec(text);
  if (!m) return '';
  const ext = m[1] === 'image/png' ? '.png' : m[1] === 'image/webp' ? '.webp' : '.jpg';
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], baseName + ext);
  return photoFolder_(stationCode).createFile(blob).getUrl();
}

function photoFolder_(stationCode) {
  const rootId = getProp_('PHOTO_FOLDER_ID');
  if (!rootId) throw new Error('尚未設定指令碼屬性 PHOTO_FOLDER_ID，請聯絡系統管理員');
  const root = DriveApp.getFolderById(rootId);
  return childFolder_(childFolder_(root, stationCode), fmt_(new Date(), 'yyyy-MM'));
}

function childFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function nextIssueSeq_(now) {
  const prefix = 'ISS-' + fmt_(now, 'yyyyMM') + '-';
  let max = 0;
  readTable_('異常追蹤').rows.forEach(r => {
    const id = String(r['異常編號'] || '');
    if (id.indexOf(prefix) === 0) max = Math.max(max, Number(id.slice(prefix.length)) || 0);
  });
  return max + 1;
}

/** 高風險異常立即寄信給站長與負責區域的安保專員；沒有 Email 時略過。 */
function notifyHighRisk_(station, issueRows, reporter, sourceLabel) {
  const high = issueRows.filter(r => r['風險等級'] === '高');
  if (!high.length) return false;
  const officer = officers_().find(o => o.on && o.regions.indexOf(station.region) >= 0);
  const to = [station.email, officer && officer.email].filter(isEmail_);
  if (!to.length || MailApp.getRemainingDailyQuota() < to.length) return false;
  const items = items_();
  const lines = high.map(r => {
    const it = items.find(i => i.code === r['項目代碼']) || {};
    return `・${r['項目代碼']} ${it.text || ''}\n  說明：${r['異常說明']}\n  改善期限：${r['改善期限']}` + (r['異常照片'] ? `\n  照片：${r['異常照片']}` : '');
  });
  MailApp.sendEmail({
    to: to.join(','),
    subject: `【高風險異常】${station.name} ${high.length} 項`,
    body: `${station.region}${station.group} ${station.name}（${sourceLabel === 'SL-H015' ? '查核人' : '填報人'}：${reporter}）在${sourceLabel === 'SL-H015' ? ' SL-H015 季查核' : '自主檢查'}中發現高風險異常：\n\n${lines.join('\n\n')}\n\n請於改善期限內處置。此信由加油站巡檢系統自動寄出。`
  });
  return true;
}


/* ============================== SL-H015 季查核（安保專員） ============================== */

/** 查核表單：負責區域的站點（含本季是否已查核）與啟用中的 SL-H015 項目。 */
function getAuditForm(token) {
  const sess = session_(token, ['安保專員']);
  const stations = scopeStations_(sess);
  const codes = {};
  stations.forEach(s => { codes[s.code] = true; });
  const audits = auditSummary_(codes);
  const items = auditItems_();
  if (!items.length) throw new Error('試算表「檢查項目」沒有啟用中的 SL-H015 項目');
  return {
    today: today_(),
    quarter: quarterLabel_(today_()),
    version: items[0].version || '',
    stations: stations.map(s => Object.assign(publicStation_(s), { auditedThisQuarter: !!audits.quarter[s.code], lastAudit: audits.last[s.code] || null })),
    items: items.map(i => ({ code: i.code, cat: i.cat, text: i.text, law: i.law, pts: i.pts, risk: i.risk, other: isOtherItem_(i) })),
    riskDays: riskDays_()
  };
}

/**
 * payload = { station, answers: [{ code, result: 'ok'|'ng'|'na', note, other, photo }] }
 * 計分：查核得分 ＝ 符合項目配分合計 ÷（滿分 − 不適用項目配分合計）× 100，取小數一位。
 */
function submitAudit(token, payload) {
  const sess = session_(token, ['安保專員']);
  payload = payload || {};
  const station = scopeStations_(sess).find(s => s.code === String(payload.station || ''));
  if (!station) throw new Error('請選擇負責區域內的站點');
  const items = auditItems_();
  const ans = {};
  (payload.answers || []).forEach(a => { if (a && a.code) ans[a.code] = a; });

  const missing = items.filter(i => !ans[i.code] || !AUDIT_LABEL[ans[i.code].result]);
  if (missing.length) throw new Error('還有 ' + missing.length + ' 項未填寫');
  const noOther = items.filter(i => isOtherItem_(i) && ans[i.code].result !== 'na' && !clean_(ans[i.code].other, 100));
  if (noOther.length) throw new Error('「其他」項目判定為符合或不符合時，請填寫項目內容');

  const score = auditScore_(items, ans);
  const now = new Date();
  const today = today_();
  const stamp = fmt_(now, 'yyyy-MM-dd HH:mm:ss');
  const recordId = 'AU-' + fmt_(now, 'yyyyMMdd-HHmmss') + '-' + station.code;
  const days = riskDays_();

  const photos = {};
  items.forEach(i => {
    const a = ans[i.code];
    if (a.result === 'ng' && a.photo) photos[i.code] = savePhoto_(a.photo, station.code, recordId + '_' + i.code);
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  const issueRows = [];
  try {
    let seq = nextIssueSeq_(now);
    const details = items.map(i => {
      const a = ans[i.code];
      const note = clean_(a.note, 500);
      const other = isOtherItem_(i) ? clean_(a.other, 100) : '';
      if (a.result === 'ng') {
        const subject = other || (i.text.length > 60 ? i.text.slice(0, 60) + '…' : i.text);
        issueRows.push({
          '異常編號': 'ISS-' + fmt_(now, 'yyyyMM') + '-' + pad_(seq++, 3), '紀錄編號': recordId, '站點代碼': station.code,
          '來源': 'SL-H015', '項目代碼': i.code, '異常說明': note ? (other ? other + '：' + note : note) : subject, '異常照片': photos[i.code] || '',
          '風險等級': i.risk || riskForPoints_(i.pts), '狀態': '待處理', '建立日期': today,
          '改善期限': addDays_(today, days[i.risk || riskForPoints_(i.pts)] || 7), '處置說明': '', '改善後照片': '',
          '結案人': '', '結案日期': '',
          '處置紀錄': JSON.stringify([{ t: stamp, who: sess.name + '（安保專員）', text: 'SL-H015 季查核不符合' + (photos[i.code] ? '，附照片 1 張' : '') }])
        });
      }
      return {
        '紀錄編號': recordId, '項目代碼': i.code, '結果': AUDIT_LABEL[a.result],
        '得分': a.result === 'ok' ? i.pts : (a.result === 'ng' ? 0 : ''),
        '說明': (other ? '其他：' + other : '') + (other && note ? '；' : '') + note,
        '照片連結': photos[i.code] || ''
      };
    });
    appendObjects_('檢查紀錄', [{
      '紀錄編號': recordId, '表單類型': AUDIT_FORM, '版本': items[0].version || '', '站點代碼': station.code, '頻率': '每季',
      '檢查日期': today, '填報人': sess.name, '送出時間': stamp, '異常數': issueRows.length,
      '得分': score.score === null ? '' : score.score, '站長確認時間': '', '站長確認人': ''
    }]);
    appendObjects_('檢查明細', details);
    appendObjects_('異常追蹤', issueRows);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  let mailed = false;
  try { mailed = notifyHighRisk_(station, issueRows, sess.name, 'SL-H015'); } catch (e) { console.warn('通知寄送失敗：' + e); }
  return Object.assign({ recordId: recordId, station: publicStation_(station), date: today, mailed: mailed, high: issueRows.filter(r => r['風險等級'] === '高').length }, score);
}

/** 站長確認查核結果（在安保專員的手機上輸入姓名），取代紙本簽名。 */
function confirmAudit(token, recordId, name) {
  const sess = session_(token, ['安保專員', '系統管理員']);
  name = clean_(name, 30);
  if (!name) throw new Error('請輸入站長姓名');
  return withLock_(() => {
    const rec = auditRecord_(sess, recordId);
    if (rec['站長確認時間']) throw new Error('這次查核已經確認過了');
    const stamp = fmt_(new Date(), 'yyyy-MM-dd HH:mm:ss');
    updateRow_('檢查紀錄', rec._row, { '站長確認時間': stamp, '站長確認人': name });
    return { recordId: recordId, confirmedAt: stamp, confirmedBy: name };
  });
}

/** 負責區域內最近的查核紀錄。 */
function listAudits(token) {
  const sess = session_(token, ['安保專員', '系統管理員']);
  const codes = {};
  scopeStations_(sess).forEach(s => { codes[s.code] = true; });
  return auditSummary_(codes).recent;
}

/**
 * 匯出一次查核為 .xlsx（格式比照 SL-H015 紙本表單）。
 * 做法：建立暫存試算表 → 以擁有者權限匯出 xlsx → 刪除暫存檔 → 回傳 base64 給前端下載。
 */
function exportAudit(token, recordId) {
  const sess = session_(token, ['安保專員', '系統管理員']);
  const rec = auditRecord_(sess, recordId);
  const station = stations_().find(s => s.code === String(rec['站點代碼'])) || { name: rec['站點代碼'], region: '', group: '' };
  const byCode = {};
  items_().forEach(i => { byCode[i.code] = i; });
  const details = readTable_('檢查明細').rows.filter(r => String(r['紀錄編號']) === recordId);
  const date = dateStr_(rec['檢查日期']);
  const confirmed = rec['站長確認時間'] ? `${rec['站長確認人']}（${String(fmtAny_(rec['站長確認時間'])).slice(0, 16)}）` : '尚未確認';

  let pass = 0, total = 0, na = 0;
  const body = details.map(d => {
    const it = byCode[String(d['項目代碼'])] || { code: d['項目代碼'], cat: '', text: '', law: '', pts: 0 };
    const pts = Number(it.pts) || 0;
    total += pts;
    if (d['結果'] === '符合') pass += pts;
    if (d['結果'] === '不適用') na += pts;
    const desc = String(d['說明'] || '');
    return [String(it.code).replace('H015-', ''), it.cat, it.text + (desc ? '\n（' + desc + '）' : ''), pts, d['結果'], d['結果'] === '不適用' ? '—' : (d['得分'] === '' ? '' : Number(d['得分'])), it.law || ''];
  });
  const base = total - na;
  const score = rec['得分'] === '' ? (base > 0 ? Math.round(pass / base * 1000) / 10 : '') : Number(rec['得分']);

  const title = '加油站安全衛生管理及營運設備查核表' + (rec['版本'] ? '（SL-H015 ' + rec['版本'] + ' 版）' : '');
  const rows = [
    [title, '', '', '', '', '', ''],
    ['站別：' + station.name + '（' + station.region + station.group + '）', '', '', '查核日期：' + date, '', '', ''],
    ['查核人：' + rec['填報人'], '', '', '站長確認：' + confirmed, '', '', ''],
    ['項次', '類別', '檢查項目', '配分', '查核情況', '得分', '法令規定']
  ].concat(body).concat([
    ['', '', '符合項目得分合計', '', '', pass, ''],
    ['', '', '滿分基準（' + total + ' 分扣除不適用 ' + na + ' 分）', '', '', base, ''],
    ['', '', '查核得分（百分制）', '', '', score, ''],
    ['備註：第一種危險場所：在加油機內部自其基礎向上至1.22公尺範圍內，及離加油機1.22公尺之四周自該基礎起向下至地面及向上至46公分處空間；第二種危險場所：在屋外其距加油機(即泵)外殼6公尺或裝有管路之油槽3公尺之四周範圍內至地面向上至46公分處。', '', '', '', '', '', '']
  ]);

  const filename = 'SL-H015查核表_' + station.name + '_' + date.split('-').join('') + '.xlsx';
  const tmp = SpreadsheetApp.create('匯出暫存_' + filename);
  const file = DriveApp.getFileById(tmp.getId());
  try {
    const folderId = getProp_('EXPORT_FOLDER_ID');
    if (folderId) file.moveTo(DriveApp.getFolderById(folderId));
    const sh = tmp.getSheets()[0];
    sh.setName('SL-H015查核表');
    const n = rows.length, head = 4, last = head + body.length;
    sh.getRange(1, 1, n, 7).setValues(rows).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
    sh.getRange(1, 1, 1, 7).merge().setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
    [2, 3].forEach(r => { sh.getRange(r, 1, 1, 3).merge(); sh.getRange(r, 4, 1, 4).merge(); });
    sh.getRange(head, 1, 1, 7).setFontWeight('bold').setBackground('#e2efe6').setHorizontalAlignment('center');
    sh.getRange(head, 1, last - head + 4, 7).setBorder(true, true, true, true, true, true);
    sh.getRange(head + 1, 3, body.length, 1).setWrap(true);
    sh.getRange(head + 1, 7, body.length, 1).setWrap(true);
    sh.getRange(head + 1, 4, body.length + 3, 3).setHorizontalAlignment('center');
    sh.getRange(last + 1, 3, 3, 1).setFontWeight('bold').setHorizontalAlignment('right');
    sh.getRange(last + 3, 6).setFontWeight('bold').setFontSize(12);
    sh.getRange(n, 1, 1, 7).merge().setWrap(true).setFontSize(9);
    [48, 96, 380, 48, 72, 48, 220].forEach((w, k) => sh.setColumnWidth(k + 1, w));
    sh.setFrozenRows(head);
    SpreadsheetApp.flush();
    const blob = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx', {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
    });
    if (blob.getResponseCode() !== 200) throw new Error('匯出失敗（' + blob.getResponseCode() + '），請稍後再試');
    return { filename: filename, data: Utilities.base64Encode(blob.getBlob().getBytes()) };
  } finally {
    try { file.setTrashed(true); } catch (e) { console.warn('暫存檔刪除失敗：' + e); }
  }
}

/** 第一次部署或更新授權範圍後，在編輯器執行一次這個函式完成授權。 */
function authorizeOnce() {
  ss_().getName();
  DriveApp.getRootFolder();
  MailApp.getRemainingDailyQuota();
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  return '授權完成';
}

function auditItems_() { return items_().filter(i => i.on && i.form === AUDIT_FORM); }
function isOtherItem_(i) { return String(i.text).indexOf('其他') === 0; }
function riskForPoints_(pts) { return pts >= 3 ? '高' : (pts >= 2 ? '中' : '低'); }

function auditScore_(items, ans) {
  let pass = 0, total = 0, na = 0, ok = 0, ng = 0, nac = 0;
  items.forEach(i => {
    const pts = Number(i.pts) || 0, r = ans[i.code].result;
    total += pts;
    if (r === 'ok') { pass += pts; ok++; }
    else if (r === 'na') { na += pts; nac++; }
    else ng++;
  });
  const base = total - na;
  return { pass: pass, total: total, na: na, base: base, okCount: ok, ngCount: ng, naCount: nac, score: base > 0 ? Math.round(pass / base * 1000) / 10 : null };
}

function auditRecord_(sess, recordId) {
  const rec = readTable_('檢查紀錄').rows.find(r => String(r['紀錄編號']) === String(recordId) && String(r['表單類型']) === AUDIT_FORM);
  if (!rec) throw new Error('找不到查核紀錄 ' + recordId);
  if (!scopeStations_(sess).some(s => s.code === String(rec['站點代碼']))) throw new Error('沒有權限查看這筆查核');
  return rec;
}

/** 依站點整理查核紀錄：最近一次、本季是否已查核、最近清單。 */
function auditSummary_(codes) {
  const q = quarterLabel_(today_());
  const names = {};
  stations_().forEach(s => { names[s.code] = s.name; });
  const recs = readTable_('檢查紀錄').rows
    .filter(r => String(r['表單類型']) === AUDIT_FORM && codes[String(r['站點代碼'])])
    .map(r => ({
      recordId: String(r['紀錄編號']), station: String(r['站點代碼']), stationName: names[String(r['站點代碼'])] || '',
      date: dateStr_(r['檢查日期']), submitted: String(fmtAny_(r['送出時間'])), auditor: String(r['填報人'] || ''),
      score: r['得分'] === '' ? null : Number(r['得分']), ng: Number(r['異常數']) || 0,
      confirmedBy: String(r['站長確認人'] || ''), confirmed: !!r['站長確認時間']
    }))
    .sort((a, b) => (b.submitted > a.submitted ? 1 : -1));
  const last = {}, quarter = {};
  recs.forEach(r => {
    if (!last[r.station]) last[r.station] = { date: r.date, score: r.score, recordId: r.recordId };
    if (quarterLabel_(r.date) === q) quarter[r.station] = true;
  });
  return { recent: recs, last: last, quarter: quarter };
}

function quarterLabel_(ymd) {
  if (!ymd) return '';
  return ymd.slice(0, 4) + ' 年第 ' + Math.ceil(Number(ymd.slice(5, 7)) / 3) + ' 季';
}

/* ============================== 異常（第一階段：唯讀） ============================== */

function listIssues(token) {
  const sess = session_(token);
  const codes = {};
  scopeStations_(sess).forEach(s => { codes[s.code] = true; });
  return issuesInScope_(codes);
}

function issuesInScope_(codes) {
  const today = today_();
  const names = {};
  stations_().forEach(s => { names[s.code] = s.name; });
  return readTable_('異常追蹤').rows
    .filter(r => codes[String(r['站點代碼'])])
    .map(r => {
      const status = String(r['狀態'] || '待處理');
      const due = dateStr_(r['改善期限']);
      let log = [];
      try { log = JSON.parse(r['處置紀錄'] || '[]'); } catch (e) { log = []; }
      return {
        id: String(r['異常編號']), station: String(r['站點代碼']), stationName: names[String(r['站點代碼'])] || '',
        src: String(r['來源'] || ''), code: String(r['項目代碼'] || ''), title: String(r['異常說明'] || ''),
        risk: String(r['風險等級'] || ''), status: status, created: dateStr_(r['建立日期']), due: due,
        closed: dateStr_(r['結案日期']), photo: String(r['異常照片'] || ''), log: log,
        overdue: status !== '已結案' && !!due && due < today
      };
    });
}

/* ============================== 管理設定（系統管理員） ============================== */

function adminData(token) {
  session_(token, ['系統管理員']);
  return { stations: stations_(), officers: officers_(), items: items_() };
}

function adminSaveStation(token, st) {
  session_(token, ['系統管理員']);
  st = st || {};
  const region = String(st.region || '');
  if (REGIONS.indexOf(region) < 0) throw new Error('區域必須是北區、中區或南區');
  const name = clean_(st.name, 20), mgr = clean_(st.mgr, 30), group = clean_(st.group, 20), email = clean_(st.email, 100);
  if (!name) throw new Error('請輸入站名');
  if (!mgr) throw new Error('請輸入站長姓名');
  if (email && !isEmail_(email)) throw new Error('站長 Email 格式不正確');
  return withLock_(() => {
    const all = stations_();
    if (all.some(s => s.name === name && s.code !== st.code)) throw new Error('已有同名站點「' + name + '」');
    if (st.code) {
      const cur = all.find(s => s.code === st.code);
      if (!cur) throw new Error('找不到站點 ' + st.code);
      updateRow_('站點', cur._row, { '站名': name, '區域': region, '組別': group, '站長': mgr, '站長Email': email });
      return st.code;
    }
    const prefix = { '北區': 'N', '中區': 'C', '南區': 'S' }[region];
    let max = 0;
    all.forEach(s => { if (s.code.indexOf(prefix) === 0) max = Math.max(max, Number(s.code.slice(1)) || 0); });
    const code = prefix + pad_(max + 1, 2);
    appendObjects_('站點', [{ '站點代碼': code, '站名': name, '區域': region, '組別': group, '站長': mgr, '站長Email': email, '啟用': '是', '備註': '' }]);
    return code;
  });
}

function adminSaveOfficer(token, o) {
  session_(token, ['系統管理員']);
  o = o || {};
  const name = clean_(o.name, 30), email = clean_(o.email, 100);
  const regions = String(o.regions || '').split(/[、,，\s]+/).filter(r => REGIONS.indexOf(r) >= 0);
  if (!name) throw new Error('請輸入姓名');
  if (!regions.length) throw new Error('請選擇負責區域');
  if (email && !isEmail_(email)) throw new Error('Email 格式不正確');
  return withLock_(() => {
    const cur = officers_().find(x => x.name === (o.originalName || name));
    const row = { '姓名': name, '角色': '安保專員', '負責區域': regions.join('、'), 'Email': email };
    if (cur) updateRow_('人員', cur._row, row);
    else appendObjects_('人員', [Object.assign(row, { '啟用': '是', '備註': '' })]);
    return name;
  });
}

function adminSaveItem(token, it) {
  session_(token, ['系統管理員']);
  it = it || {};
  const freq = String(it.freq || '');
  const prefix = { '每日': 'D', '每週': 'W', '每月': 'M' }[freq];
  if (!prefix) throw new Error('頻率必須是每日、每週或每月');
  const risk = String(it.risk || '');
  if (['高', '中', '低'].indexOf(risk) < 0) throw new Error('風險等級必須是高、中或低');
  const cat = clean_(it.cat, 20), text = clean_(it.text, 200), law = clean_(it.law, 100), std = clean_(it.std, 100);
  if (!cat || !text) throw new Error('請輸入類別與檢查內容');
  return withLock_(() => {
    const all = items_();
    if (it.code) {
      const cur = all.find(i => i.code === it.code && i.form === '自主檢查');
      if (!cur) throw new Error('只能修改自主檢查項目');
      updateRow_('檢查項目', cur._row, { '類別': cat, '檢查內容': text, '判定標準': std, '法規依據': law, '風險等級': risk });
      return it.code;
    }
    let max = 0;
    all.forEach(i => { if (i.code.indexOf('SC-' + prefix) === 0) max = Math.max(max, Number(i.code.slice(4)) || 0); });
    const code = 'SC-' + prefix + pad_(max + 1, 2);
    const order = all.reduce((m, i) => Math.max(m, i.order), 0) + 1;
    appendObjects_('檢查項目', [{
      '項目代碼': code, '表單類型': '自主檢查', '版本': 'v1', '頻率': freq, '類別': cat, '檢查內容': text,
      '判定標準': std, '法規依據': law || '公司規定', '風險等級': risk, '配分': '', '排序': order, '啟用': '是'
    }]);
    return code;
  });
}

/** kind：stations | officers | items；key：站點代碼、姓名或項目代碼 */
function adminToggle(token, kind, key) {
  session_(token, ['系統管理員']);
  return withLock_(() => {
    const map = {
      stations: ['站點', stations_(), x => x.code],
      officers: ['人員', officers_(), x => x.name],
      items: ['檢查項目', items_(), x => x.code]
    }[kind];
    if (!map) throw new Error('未知的資料類型');
    const cur = map[1].find(x => map[2](x) === key);
    if (!cur) throw new Error('找不到 ' + key);
    updateRow_(map[0], cur._row, { '啟用': cur.on ? '否' : '是' });
    return !cur.on;
  });
}

function adminCheckData(token) {
  session_(token, ['系統管理員']);
  return checkData_();
}

/* ============================== 資料檢查 ============================== */

function checkData_() {
  const out = [];
  const add = (level, sheet, msg) => out.push({ level: level, sheet: sheet, msg: msg });
  const ss = ss_();
  Object.keys(REQUIRED_HEADERS).forEach(name => {
    const sh = sheetOrNull_(ss, name);
    if (!sh) { add('錯誤', name, '找不到這個工作表（目前連到「' + ss.getName() + '」）'); return; }
    const headers = headers_(sh);
    const miss = REQUIRED_HEADERS[name].filter(h => headers.indexOf(h) < 0);
    if (miss.length) add('錯誤', name, '缺少欄位：' + miss.join('、') + '（可用選單「巡檢系統 › 補上缺少的欄位」）');
  });
  if (out.length) return out;

  const st = stations_();
  const seen = {}, seenName = {};
  st.forEach(s => {
    if (seen[s.code]) add('錯誤', '站點', '站點代碼重複：' + s.code);
    seen[s.code] = true;
    if (seenName[s.name]) add('提醒', '站點', '站名重複：' + s.name);
    seenName[s.name] = true;
    if (REGIONS.indexOf(s.region) < 0) add('錯誤', '站點', s.code + ' 區域「' + s.region + '」不是北區、中區或南區');
    if (!s.mgr) add('提醒', '站點', s.code + ' ' + s.name + ' 沒有站長');
    if (s.email && !isEmail_(s.email)) add('錯誤', '站點', s.code + ' 站長 Email 格式不正確');
  });
  const noEmail = st.filter(s => s.on && !s.email).length;
  if (noEmail) add('提醒', '站點', noEmail + ' 站尚未填站長 Email，高風險異常不會寄信給這些站長');

  const off = officers_();
  REGIONS.forEach(r => {
    if (!off.some(o => o.on && o.regions.indexOf(r) >= 0)) add('錯誤', '人員', r + '沒有啟用中的安保專員');
  });
  off.forEach(o => {
    if (o.email && !isEmail_(o.email)) add('錯誤', '人員', o.name + ' Email 格式不正確');
    if (!o.email) add('提醒', '人員', o.name + ' 尚未填 Email');
  });

  const items = items_();
  const seenItem = {};
  items.forEach(i => {
    if (seenItem[i.code]) add('錯誤', '檢查項目', '項目代碼重複：' + i.code);
    seenItem[i.code] = true;
    if (['高', '中', '低'].indexOf(i.risk) < 0) add('錯誤', '檢查項目', i.code + ' 風險等級不是高、中、低');
    if (i.form === '自主檢查' && ['每日', '每週', '每月'].indexOf(i.freq) < 0) add('錯誤', '檢查項目', i.code + ' 頻率不是每日、每週、每月');
  });
  const h015 = items.filter(i => i.on && i.form === 'SL-H015查核表').reduce((s, i) => s + (i.pts || 0), 0);
  if (h015 && h015 !== 100) add('提醒', '檢查項目', 'SL-H015 啟用項目配分合計為 ' + h015 + '，不是 100');

  const s = settings_();
  const pws = ['登入密碼_站長', '登入密碼_安保專員', '登入密碼_系統管理員'].map(k => String(s[k] === undefined ? '' : s[k]).trim());
  if (pws.some(p => !p)) add('錯誤', '設定', '登入密碼有空白');
  if (new Set(pws).size !== pws.length) add('錯誤', '設定', '三組登入密碼不能相同');
  if (pws.indexOf('12345') >= 0 || pws.indexOf('98765') >= 0 || pws.indexOf('77777') >= 0) add('提醒', '設定', '仍在使用初始密碼，正式上線前請修改');
  if (!getProp_('PHOTO_FOLDER_ID')) add('錯誤', '指令碼屬性', '尚未設定 PHOTO_FOLDER_ID');
  return out;
}

/** 依 REQUIRED_HEADERS 在工作表最右邊補上缺少的欄位；不移動、不刪除既有欄位。 */
function setupSheets_() {
  const ss = ss_();
  const added = [];
  Object.keys(REQUIRED_HEADERS).forEach(name => {
    let sh = sheetOrNull_(ss, name);
    if (!sh) { sh = ss.insertSheet(name); }
    const headers = headers_(sh);
    const miss = REQUIRED_HEADERS[name].filter(h => headers.indexOf(h) < 0);
    if (miss.length) {
      const start = headers.filter(h => h).length + 1;
      sh.getRange(1, start, 1, miss.length).setValues([miss]);
      added.push(name + '：' + miss.join('、'));
    }
  });
  return added;
}

/* ============================== 資料讀取 ============================== */

function stations_() {
  return readTable_('站點').rows.map(r => ({
    code: String(r['站點代碼'] || '').trim(), name: String(r['站名'] || '').trim(),
    region: String(r['區域'] || '').trim(), group: String(r['組別'] || '').trim(),
    mgr: String(r['站長'] || '').trim(), email: String(r['站長Email'] || '').trim(),
    on: yes_(r['啟用']), note: String(r['備註'] || ''), _row: r._row
  })).filter(s => s.code);
}

function activeStations_() { return stations_().filter(s => s.on); }

function officers_() {
  return readTable_('人員').rows
    .filter(r => String(r['角色'] || '').trim() === '安保專員')
    .map(r => ({
      name: String(r['姓名'] || '').trim(),
      regions: String(r['負責區域'] || '').split(/[、,，\s]+/).filter(x => REGIONS.indexOf(x) >= 0),
      email: String(r['Email'] || '').trim(), on: yes_(r['啟用']), _row: r._row
    })).filter(o => o.name);
}

function items_() {
  return readTable_('檢查項目').rows.map(r => ({
    code: String(r['項目代碼'] || '').trim(), form: String(r['表單類型'] || '').trim(),
    version: String(r['版本'] || ''), freq: String(r['頻率'] || '').trim(), cat: String(r['類別'] || '').trim(),
    text: String(r['檢查內容'] || '').trim(), std: String(r['判定標準'] || ''), law: String(r['法規依據'] || ''),
    risk: String(r['風險等級'] || '').trim(), pts: r['配分'] === '' || r['配分'] === null ? null : Number(r['配分']),
    order: Number(r['排序']) || 0, on: yes_(r['啟用']), _row: r._row
  })).filter(i => i.code).sort((a, b) => a.order - b.order);
}

function scopeStations_(sess) {
  const all = activeStations_();
  if (sess.role === '站長') return all.filter(s => s.code === sess.station);
  if (sess.role === '安保專員') return all.filter(s => sess.regions.indexOf(s.region) >= 0);
  return all;
}

function publicStation_(s) { return { code: s.code, name: s.name, region: s.region, group: s.group, mgr: s.mgr }; }
function publicItem_(i) { return { code: i.code, freq: i.freq, cat: i.cat, text: i.text, std: i.std, law: i.law, risk: i.risk }; }

function settings_() {
  const out = {};
  readTable_('設定').rows.forEach(r => { out[String(r['參數']).trim()] = r['值']; });
  return out;
}

function riskDays_() {
  const s = settings_();
  return {
    '高': Number(s['改善期限_高風險']) || 1,
    '中': Number(s['改善期限_中風險']) || 3,
    '低': Number(s['改善期限_低風險']) || 7
  };
}

/* ============================== 試算表工具 ============================== */

function ss_() {
  const id = String(getProp_('SHEET_ID') || '').trim();
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('程式沒有連到試算表：請在「專案設定 → 指令碼屬性」新增 SHEET_ID，值為「加油站巡檢系統資料庫」網址中 /d/ 後面那一段');
  return ss;
}

/** 依名稱找工作表；忽略名稱前後多出的空白。找不到時列出目前連到的試算表與工作表，方便排查。 */
function sheet_(name) {
  const ss = ss_();
  const sh = ss.getSheetByName(name) || ss.getSheets().find(x => String(x.getName()).trim() === name);
  if (!sh) {
    const names = ss.getSheets().map(x => '「' + x.getName() + '」').join('、');
    throw new Error('找不到工作表「' + name + '」。程式目前連到試算表「' + ss.getName() + '」，裡面的工作表有：' + names +
      '。若這不是「加油站巡檢系統資料庫」，請在指令碼屬性新增 SHEET_ID');
  }
  return sh;
}

function sheetOrNull_(ss, name) {
  return ss.getSheetByName(name) || ss.getSheets().find(x => String(x.getName()).trim() === name) || null;
}

function headers_(sh) {
  const last = sh.getLastColumn();
  if (!last) return [];
  return sh.getRange(1, 1, 1, last).getValues()[0].map(h => String(h).trim());
}

function readTable_(name) {
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  if (!values.length) return { headers: [], rows: [] };
  const headers = values[0].map(h => String(h).trim());
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (r.every(c => c === '' || c === null)) continue;
    const o = { _row: i + 1 };
    headers.forEach((h, k) => { if (h) o[h] = r[k]; });
    rows.push(o);
  }
  return { headers: headers, rows: rows };
}

function appendObjects_(name, objs) {
  if (!objs || !objs.length) return;
  const sh = sheet_(name);
  const headers = headers_(sh);
  const data = objs.map(o => headers.map(h => (h && o[h] !== undefined) ? o[h] : ''));
  sh.getRange(sh.getLastRow() + 1, 1, data.length, headers.length).setValues(data);
}

function updateRow_(name, rowNum, patch) {
  const sh = sheet_(name);
  const headers = headers_(sh);
  Object.keys(patch).forEach(k => {
    const c = headers.indexOf(k);
    if (c < 0) throw new Error('工作表「' + name + '」缺少欄位「' + k + '」');
    sh.getRange(rowNum, c + 1).setValue(patch[k]);
  });
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ============================== 小工具 ============================== */

function getProp_(key) { return PropertiesService.getScriptProperties().getProperty(key); }
function fmt_(d, pattern) { return Utilities.formatDate(d, TZ, pattern); }
function today_() { return fmt_(new Date(), 'yyyy-MM-dd'); }
function weekStart_() {
  const now = new Date();
  const dow = Number(fmt_(now, 'u')); // 1＝星期一 … 7＝星期日
  return fmt_(new Date(now.getTime() - (dow - 1) * 86400000), 'yyyy-MM-dd');
}
function addDays_(ymd, n) {
  const p = ymd.split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + n, 4)); // 台灣時間中午，避免跨日
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}
function fmtAny_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return fmt_(v, 'yyyy-MM-dd HH:mm:ss');
  return v === null || v === undefined ? '' : String(v);
}
function dateStr_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return fmt_(v, 'yyyy-MM-dd');
  return String(v).slice(0, 10).replace(/\//g, '-');
}
function pad_(n, w) { return String(n).padStart(w, '0'); }
function yes_(v) { return v === true || String(v).trim() === '是' || String(v).trim().toUpperCase() === 'TRUE'; }
function isEmail_(v) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(v || '').trim()); }
function clean_(v, max) { return String(v === undefined || v === null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max); }
