/**
 * 加油站巡檢 App — 後端（第一階段）
 *
 * 角色：站長、安保專員、油品主管（各區油品主管經理，唯讀看缺失）、系統管理員。
 * 功能：密碼登入、首頁狀態、自主檢查送出（含照片）、SL-H015 查核與匯出 Excel、營運設備現場重點項目檢查與匯出 Excel、
 *       異常清單（唯讀）、管理設定、資料檢查。
 * 資料全部存在綁定的 Google 試算表；照片存在指令碼屬性 PHOTO_FOLDER_ID 指定的資料夾。
 *
 * 指令碼屬性（專案設定 → 指令碼屬性）：
 *   PHOTO_FOLDER_ID   「巡檢照片」資料夾 ID（選填；沒填時第一次上傳照片會自動建立並記住）
 *   EXPORT_FOLDER_ID  「匯出暫存」資料夾 ID（第三階段匯出 Excel 使用）
 *   APP_NAME          App 名稱，預設「加油站巡檢」
 *   SHEET_ID          建議填寫：「加油站巡檢系統資料庫」網址 /d/ 後面那一段
 */

const TZ = 'Asia/Taipei';
const SESSION_SECONDS = 6 * 60 * 60; // CacheService 上限 6 小時
const REGIONS = ['北區', '中區', '南區'];
/** 各區油品主管經理：唯讀，可看負責區域的總覽、缺失（異常）與匯出 Excel，不能填寫或修改。 */
const MANAGER_ROLE = '油品主管';
const MANAGER_ROLE_NAMES = ['油品主管', '油品主管經理'];
const RESULT_LABEL = { ok: '正常', ng: '異常', na: '不適用' };
const MAX_PHOTO_CHARS = 4 * 1024 * 1024;
const AUDIT_FORM = 'SL-H015查核表';
const AUDIT_LABEL = { ok: '符合', ng: '不符合', na: '不適用' };

/** SL-H015 各項目的預設改善建議；試算表「檢查項目」的「改善建議」欄有填時以試算表為準。 */
const DEFAULT_SUGGEST = {
  'H015-A01': '指派人員參加丙級勞工安全衛生業務主管訓練並取得證書，證書影本存站備查。',
  'H015-A02': '每班至少一名人員完成急救人員訓練，排班表標示急救人員，證書影本存站備查。',
  'H015-A03': '每班指派一名人員完成有機溶劑作業主管訓練，證書影本存站備查。',
  'H015-A04': '相關人員應持有效之道路危險物品運送人員訓練證明，到期前安排複訓。',
  'H015-B01': '補辦新進人員體格檢查及在職人員定期健康檢查，報告依規定保存。',
  'H015-B02': '依規定配置急救箱及藥品器材，每月清點補充，過期藥品立即更換。',
  'H015-B03': '補辦新進及變更職務人員 3 小時以上安全衛生教育訓練（危險物作業另加 3 小時），保存簽到表與教材。',
  'H015-B04': '安排在職人員每年 1 小時以上安全衛生教育訓練（危險物作業另加 1 小時），保存紀錄。',
  'H015-B05': '訂定或修訂安全衛生工作守則，公告並使員工知悉。',
  'H015-B06': '於明顯易見處公告受理勞工申訴之機構、人員、範圍、格式及申訴程序。',
  'H015-C01': '清除通道油漬、積水與障礙物，修補破損地面或階梯，必要處加設止滑措施。',
  'H015-C02': '安全門於工作期間不得上鎖，移除安全門及安全梯通道堆放物品。',
  'H015-C03': '更換故障燈具，並測試緊急照明可正常點亮。',
  'H015-C04': '維持主要人行道寬度 1 公尺以上，補設安全門、安全梯明顯標示。',
  'H015-C05': '空壓機傳動帶加裝護罩及接地，完成年度定期檢查並保存紀錄。',
  'H015-C06': '洗衣機、脫水機加裝覆蓋及連鎖裝置，未改善前停止使用。',
  'H015-C07': '配電盤加裝分格板（內板），並標示各開關迴路用途。',
  'H015-C08': '移除發電室、變電室、受電室內與電路無關之物品，保持淨空。',
  'H015-C09': '電氣室門口張貼「非相關工作人員不得任意進入」標示。',
  'H015-C10': '電氣設備旁張貼「不得以濕手操作開關」「非職權範圍，不得擅自操作」標示。',
  'H015-C11': '空壓機、馬達等低壓用電設備加接地，並量測接地電阻。',
  'H015-C12': '洗車機、高壓噴槍及戶外電氣設備迴路加裝漏電斷路器，每月按測試鈕確認動作。',
  'H015-C13': '營業室前張貼有機溶劑作業注意事項公告板，並落實作業前檢點。',
  'H015-D01': '營業室內之輕質揮發性小包裝油品移至規定儲存場所。',
  'H015-D02': '儲油槽區及加油區補設紅底白字金屬板警戒標誌（嚴禁煙火、熄火加油）。',
  'H015-D03': '於明顯處補正營業主體、站名、營業時間、售油種類、油品價格及供油廠商標示。',
  'H015-D04': '補設出入口方向標誌、地面白色箭頭標線及雨棚高度標誌。',
  'H015-D05': '重新漆劃儲油槽卸油區地面黃線。',
  'H015-D06': '卸油口依油品漆上不同顏色或使用不同口徑，不使用時一律加鎖。',
  'H015-D07': '移除第一種及第二種危險場所內之非防爆電氣設備，或更換為適用防爆等級設備。',
  'H015-D08': '於卸油口、雨棚柱補貼汽油、超級柴油危害標示公告板。',
  'H015-D09': '委請專業廠商檢測避雷設備，缺損處修復並保存檢測紀錄。',
  'H015-E01': '採樣瓶依危害通識規定加貼危害圖示及標示內容。',
  'H015-E02': '安全資料表放置於工作場所易取得處，逾三年者更新版本。',
  'H015-E03': '訂定危害通識計畫並製作危害物質清單，定期更新。',
  'H015-F01': '每月實施防火避難設施自行檢查，並保存檢查紀錄。',
  'H015-F02': '更換過期或外觀損傷之滅火器，建立滅火器到期清冊每月追蹤。',
  'H015-F03': '更換故障之緊急照明燈及避難指示燈，每月測試並記錄。',
  'H015-F04': '依現況更新平面配置圖，或將防火避難設施恢復至圖面位置。',
  'H015-F05': '滅火器放置於明顯易取用處，補設紅底白字（長邊 24 公分、短邊 8 公分以上）標識。',
  'H015-F06': '調整滅火器懸掛高度：18 公斤以上上端距地 1 公尺以下，未滿 18 公斤 1.5 公尺以下。',
  'H015-G01': '依能源局「加油站營運設備自行安全檢查手冊」確實填寫，紀錄保存 1 年以上。'
};
const OTHER_SUGGEST = '依現場缺失內容限期改善，完成後拍照回報。';

/* 營運設備現場重點項目檢查表（115 年度「加油（氣）站查核與輔導」紀錄表第 1～3 頁） */
const EQUIP_FORM = '營運設備檢查表';
const EQUIP_VERSION = '115';
const EQUIP_LABEL = { ok: '良好', ng: '不良', na: '無此設備' };
const EQUIP_MARK = { '良好': 'V', '不良': 'X', '無此設備': '無此設備' };
const EQUIP_TITLE = '加油站營運設備現場重點項目檢查表紀錄表';
const EQUIP_PROGRAM = EQUIP_VERSION + ' 年度「加油（氣）站查核與輔導」';
/** 預設項目：[類別, [[檢查內容, 風險等級, 量測值單位], …]]。項目代碼 EQ-類別序-項次，例如 EQ-03-12。
 *  第一次使用時寫入試算表「檢查項目」，之後以試算表為準（可停用項目、調整風險等級）。 */
const EQUIP_CATS = [
  ['一、營業場所', [
    ['營業室無儲存有輕質之揮發性小包裝油品', '高'], ['站屋、油泵島之地面無油漬', '中'], ['排水溝無浮油／油漬', '中'],
    ['油品貯藏室無其他非油品類可燃物', '中'], ['油品貯藏室防火門無變形、損壞', '中'], ['地坪（路面）完好無嚴重龜裂、坑洞', '低']]],
  ['二、電氣設備', [
    ['配電盤操作面板前無堆置雜物且開關箱關閉', '中'], ['配電盤開關接線端子接觸良好', '中'], ['配電盤各分路絕緣電阻數值符合規定', '中', 'MΩ'],
    ['無非防爆性電氣設備或器具置於第一種場所或第二種場所範圍使用', '高'], ['電氣設備箱頭、插座無損壞，線路無鬆散裸露', '高'],
    ['電氣設備漏電斷路器試鈕作用正常', '高'], ['電氣設備接地裝置無脫落', '高'], ['避雷設施接地裝置無脫落', '中'],
    ['緊急照明燈功能正常', '中'], ['油品儲藏室防爆開關及防爆燈具完好', '高']]],
  ['三、儲油設備', [
    ['卸油口蓋上鎖', '高'], ['卸油口蓋及墊圈完好', '中'], ['卸油口油品類別標識正確完好', '中'], ['卸油口盛油盤內無雜物及卸水口關閉', '中'],
    ['卸油口靜電接地棒無毀損、鍍銅無斷裂', '高'], ['儲油槽陰井蓋板完好', '中'], ['儲油槽陰井內無積水或積油', '中'], ['陰井內部及壁板完整', '低'],
    ['陰井蓋板油品類別顏色標識正確完好', '低'], ['手動量油口蓋墊圈完好、蓋妥（放蓋後上鎖）', '中'], ['陰井內油管法蘭及接頭無滲漏', '高'],
    ['陰井之油氣濃度未逾 25%LEL', '高', '%LEL'], ['沈油泵馬達等接線盒內無積水、積油', '高'], ['排氣（通氣）管防雨蓋完好', '低']]],
  ['四、加油設備（加油機）', [
    ['加油機底座穩固', '中'], ['加油機底座四周無漏油跡象', '高'], ['加油槍及橡皮管無滲漏油、無破損', '高'], ['加油槍之加滿自動跳停裝置功能良好', '高'],
    ['加油橡皮管吊管回收裝置功能良好', '低'], ['運轉無異常', '中'], ['加油機內無積水、無積油及無油氣，設施接頭無滲漏', '高'], ['加油機接地裝置無脫落', '高'],
    ['傳動皮帶鬆緊度適當【自吸式】無龜裂', '低'], ['加油機及電氣接線孔配電管接頭、防爆軟管接頭無鬆脫、預留管口密封', '高'], ['沈油泵加油機緊急遮斷閥固定', '高']]],
  ['五、油氣回收設備', [
    ['第一階段油氣回收快速接頭蓋上鎖', '中'], ['第一階段油氣回收口快速接頭蓋墊圈及操作良好', '中'], ['油氣回收槍及同軸皮管無滲漏', '中']]],
  ['六、消防設備（滅火器）', [
    ['放置定位', '中'], ['消防設備前無堆置雜物', '中'], ['滅火器及藥劑未逾有效期限', '高'], ['滅火器本體無損傷、鏽蝕', '中'], ['噴嘴口無異物堵塞', '中'],
    ['噴嘴把手靈活（加壓式滅火器）', '中'], ['橡皮管無損傷、接頭旋緊、無鏽蝕／裂痕', '中'], ['封條完好', '低'], ['蓄壓式滅火器壓力指示正常（指針在綠色範圍）', '高']]],
  ['七、標示設備', [
    ['「熄火加油」及「嚴禁煙火」標誌完好', '中'], ['卸油區地面黃線標示清楚', '低'], ['車輛出入口立牆標誌及地面標線完好', '低'], ['雨棚高度標誌完好', '低'],
    ['燈柱（框架）及基座完好、無鏽損', '低'], ['接線蓋板完好', '中']]],
  ['八、靜電（電氣）接地設備', [
    ['收油口靜電接地電阻 ＜50Ω', '高', 'Ω'], ['加油機接地電阻 ＜25Ω', '高', 'Ω'], ['配電盤接地電阻 ＜50Ω', '高', 'Ω'],
    ['避雷針接地電阻 ＜10Ω', '中', 'Ω'], ['空氣壓縮機接地電阻 ＜50Ω', '中', 'Ω']]],
  ['九、空氣壓縮機', [
    ['運轉無異狀', '低'], ['接頭無漏氣', '低'], ['壓力計壓力指示正常', '中'], ['曲軸箱之機油油位適當', '低'], ['空壓機接地裝置無脫落', '中'],
    ['螺栓及管線接頭無鬆動', '低'], ['皮帶鬆緊度適當、無龜裂', '低'], ['壓力釋放並啟動洩壓閥，動作正常', '高'], ['安全閥之作用正常', '高'], ['防護網完好', '中']]],
  ['十、洗車機', [['運轉無異狀', '低'], ['漏電斷路器測試正常', '高'], ['緊急停止裝置測試正常', '高']]],
  ['十一、發電機', [['冷卻水、潤滑油及電瓶液位正常', '低'], ['燃油量充足', '低']]],
  ['十二、特別規定', [
    ['每日盤查各儲油槽油品存量、並作成紀錄', '中'], ['每月盤查儲油槽存量及自主檢測測漏管油氣濃度（LEL%）並紀錄', '中'],
    ['每年應依消防法規定，由領有證照之專業人員執行消防檢查，並檢附檢查報告，加油站自行安全檢查負責人應檢查相關紀錄與報告是否保存完好', '中']]]
];
const EQUIP_UNIT = {};
EQUIP_CATS.forEach((c, ci) => c[1].forEach((x, k) => { if (x[2]) EQUIP_UNIT[equipCode_(ci + 1, k + 1)] = x[2]; }));
function equipCode_(cat, no) { return 'EQ-' + pad_(cat, 2) + '-' + pad_(no, 2); }

const REQUIRED_HEADERS = {
  '站點': ['站點代碼', '站名', '區域', '組別', '站長', '站長Email', '啟用', '備註'],
  '人員': ['姓名', '角色', '負責區域', 'Email', '啟用', '備註'],
  '檢查項目': ['項目代碼', '表單類型', '版本', '頻率', '類別', '檢查內容', '判定標準', '法規依據', '風險等級', '配分', '排序', '啟用'],
  '檢查紀錄': ['紀錄編號', '表單類型', '版本', '站點代碼', '頻率', '檢查日期', '填報人', '送出時間', '異常數', '得分', '站長確認時間', '站長確認人'],
  '檢查明細': ['紀錄編號', '項目代碼', '結果', '得分', '說明', '照片連結', '改善建議'],
  '異常追蹤': ['異常編號', '紀錄編號', '站點代碼', '來源', '項目代碼', '異常說明', '異常照片', '風險等級', '狀態', '建立日期', '改善期限', '處置說明', '改善後照片', '結案人', '結案日期', '處置紀錄', '改善建議'],
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
  if (role === '安保專員' || role === MANAGER_ROLE) {
    const all = activeStations_();
    out.officers = (role === MANAGER_ROLE ? managers_() : officers_()).filter(o => o.on).map(o => ({
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
  } else if (role === MANAGER_ROLE) {
    const m = managers_().find(x => x.on && x.name === String(req.officer || ''));
    if (!m) throw new Error('請選擇你的姓名');
    sess.name = m.name; sess.regions = m.regions;
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
  const map = { '登入密碼_站長': '站長', '登入密碼_安保專員': '安保專員', '登入密碼_系統管理員': '系統管理員', '登入密碼_油品主管': MANAGER_ROLE, '登入密碼_油品主管經理': MANAGER_ROLE };
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
  const equips = equipSummary_(codes);
  return {
    session: sess,
    today: today,
    stations: stations.map(s => Object.assign(publicStation_(s), status[s.code], { lastAudit: audits.last[s.code] || null, auditedThisQuarter: !!audits.quarter[s.code], lastEquip: equips.last[s.code] || null })),
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
    if (a.result === 'ng' && a.photo) photos[i.code] = photoValue_(a.photo, station.code, recordId + '_' + i.code);
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

/**
 * 拍照後立即上傳：回傳雲端硬碟連結，畫面草稿只保存連結，切換分頁或重新整理都不會遺失照片。
 * 站長只能上傳自己的站；安保專員只能上傳負責區域的站。
 */
function uploadPhoto(token, stationCode, itemCode, dataUrl) {
  const sess = session_(token, ['站長', '安保專員']);
  const st = scopeStations_(sess).find(s => s.code === String(stationCode || ''));
  if (!st) throw new Error('沒有權限上傳這個站點的照片');
  const name = (sess.role === '站長' ? 'CK' : 'AU') + '-' + fmt_(new Date(), 'yyyyMMdd-HHmmss') + '-' + st.code + '_' + clean_(itemCode, 20);
  const url = savePhoto_(dataUrl, st.code, name);
  if (!url) throw new Error('照片格式不正確，請重新拍攝');
  return { url: url };
}

/** 送出時的照片欄位：已上傳的雲端硬碟連結直接使用；舊版畫面傳來的 data URL 則先存檔。 */
function photoValue_(v, stationCode, baseName) {
  const text = String(v || '');
  if (!text) return '';
  if (text.indexOf('https://drive.google.com/') === 0) return text;
  if (text.indexOf('data:image/') === 0) return savePhoto_(text, stationCode, baseName);
  return '';
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
  return childFolder_(childFolder_(photoRoot_(), stationCode), fmt_(new Date(), 'yyyy-MM'));
}

/**
 * 照片根資料夾：優先用指令碼屬性 PHOTO_FOLDER_ID；沒有設定或資料夾已不存在時，
 * 自動在試算表所在的資料夾建立（或沿用）「巡檢照片」資料夾，並把 ID 記到 PHOTO_FOLDER_ID。
 */
function photoRoot_() {
  const props = PropertiesService.getScriptProperties();
  const id = String(props.getProperty('PHOTO_FOLDER_ID') || '').trim();
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { console.warn('PHOTO_FOLDER_ID 無法開啟，改為自動建立：' + e); }
  }
  let parent;
  try {
    const parents = DriveApp.getFileById(ss_().getId()).getParents();
    parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  } catch (e) { parent = DriveApp.getRootFolder(); }
  const folder = childFolder_(parent, '巡檢照片');
  props.setProperty('PHOTO_FOLDER_ID', folder.getId());
  return folder;
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
    body: `${station.region}${station.group} ${station.name}（${{ '自主檢查': '填報人', 'SL-H015': '查核人' }[sourceLabel] || '檢查人'}：${reporter}）在${sourceLabel === 'SL-H015' ? ' SL-H015 季查核' : sourceLabel}中發現高風險異常：\n\n${lines.join('\n\n')}\n\n請於改善期限內處置。此信由加油站巡檢系統自動寄出。`
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
    items: items.map(i => ({ code: i.code, cat: i.cat, text: i.text, law: i.law, pts: i.pts, risk: i.risk, other: isOtherItem_(i), suggest: suggestFor_(i) })),
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

  ensureColumns_('檢查紀錄', ['站長確認人']);
  ensureColumns_('檢查明細', ['改善建議']);
  ensureColumns_('異常追蹤', ['改善建議']);
  const score = auditScore_(items, ans);
  const now = new Date();
  const today = today_();
  const stamp = fmt_(now, 'yyyy-MM-dd HH:mm:ss');
  const recordId = 'AU-' + fmt_(now, 'yyyyMMdd-HHmmss') + '-' + station.code;
  const days = riskDays_();

  const photos = {};
  items.forEach(i => {
    const a = ans[i.code];
    if (a.result === 'ng' && a.photo) photos[i.code] = photoValue_(a.photo, station.code, recordId + '_' + i.code);
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
      const sugg = a.result === 'ng' ? (clean_(a.suggest, 500) || suggestFor_(i)) : '';
      if (a.result === 'ng') {
        const subject = other || (i.text.length > 60 ? i.text.slice(0, 60) + '…' : i.text);
        issueRows.push({
          '異常編號': 'ISS-' + fmt_(now, 'yyyyMM') + '-' + pad_(seq++, 3), '紀錄編號': recordId, '站點代碼': station.code,
          '來源': 'SL-H015', '項目代碼': i.code, '異常說明': note ? (other ? other + '：' + note : note) : subject, '異常照片': photos[i.code] || '',
          '風險等級': i.risk || riskForPoints_(i.pts), '狀態': '待處理', '建立日期': today,
          '改善期限': addDays_(today, days[i.risk || riskForPoints_(i.pts)] || 7), '處置說明': '', '改善後照片': '',
          '結案人': '', '結案日期': '',
          '處置紀錄': JSON.stringify([{ t: stamp, who: sess.name + '（安保專員）', text: 'SL-H015 季查核不符合' + (photos[i.code] ? '，附照片 1 張' : '') }]),
          '改善建議': sugg
        });
      }
      return {
        '紀錄編號': recordId, '項目代碼': i.code, '結果': AUDIT_LABEL[a.result],
        '得分': a.result === 'ok' ? i.pts : (a.result === 'ng' ? 0 : ''),
        '說明': (other ? '其他：' + other : '') + (other && note ? '；' : '') + note,
        '照片連結': photos[i.code] || '', '改善建議': sugg
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
  const sess = session_(token, ['安保專員', '系統管理員', MANAGER_ROLE]);
  const codes = {};
  scopeStations_(sess).forEach(s => { codes[s.code] = true; });
  return auditSummary_(codes).recent;
}

/**
 * 匯出一次查核為 .xlsx。
 * mode = 'ng'：只匯出不符合項目（含照片、改善建議、改善期限）。
 * mode = 'all'：第一個工作表為完整 SL-H015 查核表（比照紙本），第二個工作表為不符合項目與照片。
 * 做法：建立暫存試算表 → 以擁有者權限匯出 xlsx → 刪除暫存檔 → 回傳 base64 給前端下載。
 */
function exportAudit(token, recordId, mode) {
  const sess = session_(token, ['安保專員', '系統管理員', MANAGER_ROLE]);
  mode = mode === 'all' ? 'all' : 'ng';
  const rec = auditRecord_(sess, recordId);
  const station = stations_().find(s => s.code === String(rec['站點代碼'])) || { name: rec['站點代碼'], region: '', group: '' };
  const byCode = {};
  items_().forEach(i => { byCode[i.code] = i; });
  const details = readTable_('檢查明細').rows.filter(r => String(r['紀錄編號']) === recordId);
  const issues = {};
  readTable_('異常追蹤').rows.filter(r => String(r['紀錄編號']) === recordId).forEach(r => { issues[String(r['項目代碼'])] = r; });
  const date = dateStr_(rec['檢查日期']);
  const confirmed = rec['站長確認時間'] ? `${rec['站長確認人']}（${String(fmtAny_(rec['站長確認時間'])).slice(0, 16)}）` : '尚未確認';

  let pass = 0, total = 0, na = 0;
  const all = details.map(d => {
    const it = byCode[String(d['項目代碼'])] || { code: String(d['項目代碼']), cat: '', text: '', law: '', pts: 0, risk: '' };
    const pts = Number(it.pts) || 0;
    total += pts;
    if (d['結果'] === '符合') pass += pts;
    if (d['結果'] === '不適用') na += pts;
    return { it: it, d: d, pts: pts, iss: issues[it.code] || null };
  });
  const base = total - na;
  const score = rec['得分'] === '' ? (base > 0 ? Math.round(pass / base * 1000) / 10 : '') : Number(rec['得分']);
  const ng = all.filter(x => x.d['結果'] === '不符合');
  const filename = 'SL-H015' + (mode === 'ng' ? '不符合改善事項_' : '查核表_') + station.name + '_' + date.split('-').join('') + '.xlsx';
  const titleBase = '加油站安全衛生管理及營運設備查核表' + (rec['版本'] ? '（SL-H015 ' + rec['版本'] + ' 版）' : '');
  const info = [
    '站別：' + station.name + '（' + station.region + station.group + '）', '查核日期：' + date,
    '查核人：' + rec['填報人'], '站長確認：' + confirmed,
    '查核得分：' + (score === '' ? '—' : score) + ' 分（符合 ' + pass + ' ÷ 基準 ' + base + ' × 100）', '不符合項目：' + ng.length + ' 項'
  ];

  return exportXlsx_(filename, tmp => {
    const first = tmp.getSheets()[0];
    if (mode === 'all') {
      writeFullSheet_(first, titleBase, info, all, pass, total, na, base, score);
      writeNgSheet_(tmp.insertSheet('不符合改善事項'), '不符合項目改善事項', info, ng);
    } else {
      writeNgSheet_(first, '不符合項目改善事項', info, ng);
    }
  });
}

/** 建立暫存試算表 → build 寫入內容 → 以擁有者權限匯出 xlsx → 刪除暫存檔，回傳 base64 給前端下載。 */
function exportXlsx_(filename, build) {
  const tmp = SpreadsheetApp.create('匯出暫存_' + filename);
  const file = DriveApp.getFileById(tmp.getId());
  try {
    const folderId = getProp_('EXPORT_FOLDER_ID');
    if (folderId) { try { file.moveTo(DriveApp.getFolderById(folderId)); } catch (e) { console.warn('EXPORT_FOLDER_ID 無法使用：' + e); } }
    build(tmp);
    SpreadsheetApp.flush();
    const res = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx', {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) throw new Error('匯出失敗（' + res.getResponseCode() + '），請稍後再試');
    return { filename: filename, data: Utilities.base64Encode(res.getBlob().getBytes()) };
  } finally {
    try { file.setTrashed(true); } catch (e) { console.warn('暫存檔刪除失敗：' + e); }
  }
}

/** 完整查核表（比照 SL-H015 紙本）。 */
function writeFullSheet_(sh, title, info, all, pass, total, na, base, score) {
  sh.setName('SL-H015查核表');
  const body = all.map(x => {
    const desc = String(x.d['說明'] || '');
    return [String(x.it.code).replace('H015-', ''), x.it.cat, x.it.text + (desc ? '\n（' + desc + '）' : ''), x.pts, x.d['結果'],
      x.d['結果'] === '不適用' ? '—' : (x.d['得分'] === '' ? '' : Number(x.d['得分'])), x.it.law || ''];
  });
  const rows = [
    [title, '', '', '', '', '', ''],
    [info[0], '', '', info[1], '', '', ''],
    [info[2], '', '', info[3], '', '', ''],
    ['項次', '類別', '檢查項目', '配分', '查核情況', '得分', '法令規定']
  ].concat(body).concat([
    ['', '', '符合項目得分合計', '', '', pass, ''],
    ['', '', '滿分基準（' + total + ' 分扣除不適用 ' + na + ' 分）', '', '', base, ''],
    ['', '', '查核得分（百分制）', '', '', score, ''],
    ['備註：第一種危險場所：在加油機內部自其基礎向上至1.22公尺範圍內，及離加油機1.22公尺之四周自該基礎起向下至地面及向上至46公分處空間；第二種危險場所：在屋外其距加油機(即泵)外殼6公尺或裝有管路之油槽3公尺之四周範圍內至地面向上至46公分處。', '', '', '', '', '', '']
  ]);
  const n = rows.length, head = 4, last = head + body.length;
  sh.getRange(1, 1, n, 7).setValues(rows).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  sh.getRange(1, 1, 1, 7).merge().setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  [2, 3].forEach(r => { sh.getRange(r, 1, 1, 3).merge(); sh.getRange(r, 4, 1, 4).merge(); });
  sh.getRange(head, 1, 1, 7).setFontWeight('bold').setBackground('#e2efe6').setHorizontalAlignment('center');
  sh.getRange(head, 1, last - head + 4, 7).setBorder(true, true, true, true, true, true);
  sh.getRange(head + 1, 3, body.length, 1).setWrap(true);
  sh.getRange(head + 1, 7, body.length, 1).setWrap(true);
  sh.getRange(head + 1, 4, body.length + 3, 3).setHorizontalAlignment('center');
  all.forEach((x, k) => { if (x.d['結果'] === '不符合') sh.getRange(head + 1 + k, 5).setFontColor('#c0352d').setFontWeight('bold'); });
  sh.getRange(last + 1, 3, 3, 1).setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange(last + 3, 6).setFontWeight('bold').setFontSize(12);
  sh.getRange(n, 1, 1, 7).merge().setWrap(true).setFontSize(9);
  [48, 96, 380, 48, 72, 48, 220].forEach((w, k) => sh.setColumnWidth(k + 1, w));
  sh.setFrozenRows(head);
}

/** 不符合項目改善事項：每項一列，含照片（嵌入儲存格上方）、改善建議與改善期限。 */
function writeNgSheet_(sh, title, info, ng, words) {
  words = Object.assign({ sheet: '不符合改善事項', ng: '不符合', empty: '本次查核沒有不符合項目' }, words || {});
  sh.setName(words.sheet);
  const head = 5, PH = 150;
  const header = ['項次', '檢查項目', words.ng + '說明', '照片', '改善建議', '風險', '改善期限', '照片連結'];
  const body = ng.map(x => {
    const iss = x.iss || {};
    return [shortCode_(x.it.code), (x.it.cat && words.cat ? x.it.cat + '：' : '') + x.it.text + (x.it.law ? '\n（' + x.it.law + '）' : ''),
      String(x.d['說明'] || '') || String(iss['異常說明'] || ''), '',
      String(x.d['改善建議'] || '') || String(iss['改善建議'] || '') || suggestFor_(x.it),
      String(iss['風險等級'] || x.it.risk || ''), dateStr_(iss['改善期限']), String(x.d['照片連結'] || '')];
  });
  const rows = [
    [title, '', '', '', '', '', '', ''],
    [info[0], '', '', '', info[1], '', '', ''],
    [info[2], '', '', '', info[3], '', '', ''],
    [info[4], '', '', '', info[5], '', '', ''],
    header
  ].concat(body.length ? body : [['', words.empty, '', '', '', '', '', '']]);
  const n = rows.length;
  sh.getRange(1, 1, n, 8).setValues(rows).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  sh.getRange(1, 1, 1, 8).merge().setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  [2, 3, 4].forEach(r => { sh.getRange(r, 1, 1, 4).merge(); sh.getRange(r, 5, 1, 4).merge(); });
  sh.getRange(head, 1, 1, 8).setFontWeight('bold').setBackground('#fbe8e6').setHorizontalAlignment('center');
  sh.getRange(head, 1, n - head + 1, 8).setBorder(true, true, true, true, true, true);
  if (body.length) {
    sh.getRange(head + 1, 2, body.length, 4).setWrap(true).setVerticalAlignment('top');
    sh.getRange(head + 1, 1, body.length, 1).setHorizontalAlignment('center');
    sh.getRange(head + 1, 6, body.length, 2).setHorizontalAlignment('center');
    sh.getRange(head + 1, 8, body.length, 1).setWrap(true).setFontSize(8);
  }
  [44, 230, 180, 210, 240, 48, 84, 120].forEach((w, k) => sh.setColumnWidth(k + 1, w));
  ng.forEach((x, k) => {
    const row = head + 1 + k, url = String(x.d['照片連結'] || '');
    sh.setRowHeight(row, url ? PH + 12 : 60);
    if (!url) { sh.getRange(row, 4).setValue('（未拍照）').setHorizontalAlignment('center'); return; }
    try {
      const img = sh.insertImage(photoBlob_(url), 4, row, 6, 6);
      const w = img.getWidth(), h = img.getHeight(), k2 = Math.min(198 / w, PH / h, 1);
      img.setWidth(Math.round(w * k2)).setHeight(Math.round(h * k2));
    } catch (e) {
      console.warn('照片嵌入失敗：' + e);
      sh.getRange(row, 4).setValue('照片無法嵌入，請開啟右側連結').setWrap(true);
    }
  });
  sh.setFrozenRows(head);
}

/** 取得照片內容：先取 Drive 縮圖（寬 600），失敗時用原檔。 */
function photoBlob_(url) {
  const m = /\/d\/([A-Za-z0-9_-]+)/.exec(url) || /[?&]id=([A-Za-z0-9_-]+)/.exec(url);
  if (!m) throw new Error('照片連結格式不正確');
  try {
    const r = UrlFetchApp.fetch('https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w600', {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
    });
    if (r.getResponseCode() === 200 && String(r.getHeaders()['Content-Type'] || r.getHeaders()['content-type'] || '').indexOf('image') === 0) return r.getBlob();
  } catch (e) { console.warn('縮圖取得失敗：' + e); }
  return DriveApp.getFileById(m[1]).getBlob();
}

/** 畫面與 Excel 用的短項次：H015-D06 → D06；EQ-03-12 → 3-12。 */
function shortCode_(code) {
  const m = /^EQ-(\d+)-(\d+)$/.exec(String(code));
  return m ? Number(m[1]) + '-' + Number(m[2]) : String(code).replace('H015-', '');
}

function suggestFor_(i) {
  if (i.suggest) return i.suggest;
  if (isOtherItem_(i)) return OTHER_SUGGEST;
  return DEFAULT_SUGGEST[i.code] || OTHER_SUGGEST;
}

/** 工作表缺少欄位時，在最右邊補上（不移動既有欄位）。 */
function ensureColumns_(name, cols) {
  const sh = sheet_(name);
  const headers = headers_(sh);
  const miss = cols.filter(c => headers.indexOf(c) < 0);
  if (miss.length) sh.getRange(1, headers.filter(h => h).length + 1, 1, miss.length).setValues([miss]);
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

/* ============================== 營運設備現場重點項目檢查（站長、安保專員） ============================== */

/** 檢查表單：可檢查的站點（站長只有自己的站）、啟用中的項目與最近紀錄。 */
function getEquipForm(token) {
  const sess = session_(token, ['站長', '安保專員']);
  ensureEquipItems_();
  const stations = scopeStations_(sess);
  const codes = {};
  stations.forEach(s => { codes[s.code] = true; });
  const sum = equipSummary_(codes);
  const items = equipItems_();
  if (!items.length) throw new Error('試算表「檢查項目」沒有啟用中的營運設備檢查項目');
  return {
    today: today_(),
    title: EQUIP_TITLE, program: EQUIP_PROGRAM, version: items[0].version || EQUIP_VERSION,
    stations: stations.map(s => Object.assign(publicStation_(s), { lastEquip: sum.last[s.code] || null })),
    items: items.map(i => ({ code: i.code, no: shortCode_(i.code), cat: i.cat, text: i.text, std: i.std, risk: i.risk || '中', unit: EQUIP_UNIT[i.code] || '' })),
    recent: sum.recent.slice(0, 10),
    riskDays: riskDays_()
  };
}

/**
 * payload = { station, answers: [{ code, result: 'ok'|'ng'|'na', note, measure, photo }] }
 * 不良項目建立異常追蹤（來源「營運設備檢查」），改善期限依項目風險等級；不計分。
 */
function submitEquip(token, payload) {
  const sess = session_(token, ['站長', '安保專員']);
  payload = payload || {};
  const station = scopeStations_(sess).find(s => s.code === String(payload.station || ''));
  if (!station) throw new Error(sess.role === '站長' ? '只能檢查自己的站點' : '請選擇負責區域內的站點');
  const items = equipItems_();
  if (!items.length) throw new Error('目前沒有啟用的營運設備檢查項目');
  const ans = {};
  (payload.answers || []).forEach(a => { if (a && a.code) ans[a.code] = a; });
  const missing = items.filter(i => !ans[i.code] || !EQUIP_LABEL[ans[i.code].result]);
  if (missing.length) throw new Error('還有 ' + missing.length + ' 項未填寫');

  ensureColumns_('檢查紀錄', ['站長確認人']);
  ensureColumns_('異常追蹤', ['改善建議']);
  const now = new Date();
  const today = today_();
  const stamp = fmt_(now, 'yyyy-MM-dd HH:mm:ss');
  let recordId = 'EI-' + fmt_(now, 'yyyyMMdd-HHmmss') + '-' + station.code;
  const days = riskDays_();
  const who = sess.name + '（' + sess.role + '）';

  const photos = {};
  items.forEach(i => {
    const a = ans[i.code];
    if (a.result === 'ng' && a.photo) photos[i.code] = photoValue_(a.photo, station.code, recordId + '_' + i.code);
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  const issueRows = [];
  const count = { ok: 0, ng: 0, na: 0 };
  try {
    recordId = uniqueRecordId_(recordId);
    let seq = nextIssueSeq_(now);
    const details = items.map(i => {
      const a = ans[i.code];
      count[a.result]++;
      const note = clean_(a.note, 500);
      const desc = [measureText_(a.measure, EQUIP_UNIT[i.code]), note].filter(x => x).join('；');
      if (a.result === 'ng') {
        const risk = ['高', '中', '低'].indexOf(i.risk) >= 0 ? i.risk : '中';
        issueRows.push({
          '異常編號': 'ISS-' + fmt_(now, 'yyyyMM') + '-' + pad_(seq++, 3), '紀錄編號': recordId, '站點代碼': station.code,
          '來源': '營運設備檢查', '項目代碼': i.code, '異常說明': i.cat.replace(/^[一二三四五六七八九十]+、/, '') + '：' + (desc ? i.text + '（' + desc + '）' : i.text),
          '異常照片': photos[i.code] || '', '風險等級': risk, '狀態': '待處理', '建立日期': today,
          '改善期限': addDays_(today, days[risk] || 7), '處置說明': '', '改善後照片': '', '結案人': '', '結案日期': '',
          '處置紀錄': JSON.stringify([{ t: stamp, who: who, text: '營運設備檢查不良' + (photos[i.code] ? '，附照片 1 張' : '') }]),
          '改善建議': ''
        });
      }
      return { '紀錄編號': recordId, '項目代碼': i.code, '結果': EQUIP_LABEL[a.result], '得分': '', '說明': desc, '照片連結': photos[i.code] || '' };
    });
    appendObjects_('檢查紀錄', [{
      '紀錄編號': recordId, '表單類型': EQUIP_FORM, '版本': items[0].version || EQUIP_VERSION, '站點代碼': station.code, '頻率': '不定期',
      '檢查日期': today, '填報人': sess.name, '送出時間': stamp, '異常數': issueRows.length, '得分': '', '站長確認時間': '', '站長確認人': ''
    }]);
    appendObjects_('檢查明細', details);
    appendObjects_('異常追蹤', issueRows);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  let mailed = false;
  try { mailed = notifyHighRisk_(station, issueRows, sess.name, '營運設備檢查'); } catch (e) { console.warn('通知寄送失敗：' + e); }
  return {
    recordId: recordId, station: publicStation_(station), date: today, total: items.length,
    okCount: count.ok, ngCount: count.ng, naCount: count.na,
    high: issueRows.filter(r => r['風險等級'] === '高').length, mailed: mailed
  };
}

/** 可查看範圍內最近的營運設備檢查紀錄。 */
function listEquips(token) {
  const sess = session_(token);
  const codes = {};
  scopeStations_(sess).forEach(s => { codes[s.code] = true; });
  return equipSummary_(codes).recent;
}

/**
 * 匯出一次營運設備檢查為 .xlsx。
 * mode = 'all'：第一個工作表比照紙本紀錄表（設備名稱、檢查項目、良好 V／不良 X／無此設備、缺失說明），第二個工作表為不良項目與照片。
 * mode = 'ng'：只匯出不良項目（含照片、改善期限）。
 */
function exportEquip(token, recordId, mode) {
  const sess = session_(token);
  mode = mode === 'ng' ? 'ng' : 'all';
  const rec = readTable_('檢查紀錄').rows.find(r => String(r['紀錄編號']) === String(recordId) && String(r['表單類型']) === EQUIP_FORM);
  if (!rec) throw new Error('找不到檢查紀錄 ' + recordId);
  if (!scopeStations_(sess).some(s => s.code === String(rec['站點代碼']))) throw new Error('沒有權限查看這筆檢查');
  const station = stations_().find(s => s.code === String(rec['站點代碼'])) || { name: rec['站點代碼'], region: '', group: '' };
  const byCode = {};
  items_().forEach(i => { byCode[i.code] = i; });
  const issues = {};
  readTable_('異常追蹤').rows.filter(r => String(r['紀錄編號']) === recordId).forEach(r => { issues[String(r['項目代碼'])] = r; });
  const all = readTable_('檢查明細').rows.filter(r => String(r['紀錄編號']) === recordId).map(d => {
    const it = byCode[String(d['項目代碼'])] || { code: String(d['項目代碼']), cat: '', text: '', law: '', risk: '' };
    return { it: it, d: d, iss: issues[it.code] || null };
  });
  const ng = all.filter(x => x.d['結果'] === '不良');
  const cnt = r => all.filter(x => x.d['結果'] === r).length;
  const date = dateStr_(rec['檢查日期']);
  const filename = (mode === 'ng' ? '營運設備不良改善事項_' : '營運設備檢查表_') + station.name + '_' + date.split('-').join('') + '.xlsx';
  const info = [
    '站別：' + station.name + '（' + station.region + station.group + '）', '檢查日期：' + date,
    '檢查人：' + rec['填報人'], '紀錄編號：' + recordId,
    '檢查結果：良好 ' + cnt('良好') + ' 項、不良 ' + ng.length + ' 項、無此設備 ' + cnt('無此設備') + ' 項', '不良項目：' + ng.length + ' 項'
  ];
  const words = { sheet: '不良改善事項', ng: '不良', empty: '本次檢查沒有不良項目', cat: true };
  return exportXlsx_(filename, tmp => {
    const first = tmp.getSheets()[0];
    if (mode === 'all') {
      writeEquipSheet_(first, info, all);
      writeNgSheet_(tmp.insertSheet('不良改善事項'), '營運設備不良項目改善事項', info, ng, words);
    } else {
      writeNgSheet_(first, '營運設備不良項目改善事項', info, ng, words);
    }
  });
}

/** 比照紙本：115 年度「加油（氣）站查核與輔導」加油站營運設備現場重點項目檢查表紀錄表。 */
function writeEquipSheet_(sh, info, all) {
  sh.setName('營運設備檢查表');
  const head = 5;
  const body = all.map(x => {
    const m = /-(\d+)$/.exec(x.it.code);
    return [x.it.cat, (m ? Number(m[1]) + '. ' : '') + x.it.text, EQUIP_MARK[x.d['結果']] || String(x.d['結果'] || ''), String(x.d['說明'] || '')];
  });
  const rows = [
    [EQUIP_PROGRAM, '', '', ''],
    [EQUIP_TITLE, '', '', ''],
    [info[0], '', info[1], ''],
    [info[2], '', info[4], ''],
    ['設備名稱', '檢查項目及結果', '良好 V\n不良 X\n無此設備', '缺失說明（備註）']
  ].concat(body);
  const n = rows.length;
  sh.getRange(1, 1, n, 4).setValues(rows).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  sh.getRange(1, 1, 1, 4).merge().setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2, 1, 1, 4).merge().setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');
  [3, 4].forEach(r => { sh.getRange(r, 1, 1, 2).merge(); sh.getRange(r, 3, 1, 2).merge(); });
  sh.getRange(head, 1, 1, 4).setFontWeight('bold').setBackground('#e2efe6').setHorizontalAlignment('center').setWrap(true);
  sh.getRange(head, 1, n - head + 1, 4).setBorder(true, true, true, true, true, true);
  if (body.length) {
    sh.getRange(head + 1, 2, body.length, 1).setWrap(true);
    sh.getRange(head + 1, 4, body.length, 1).setWrap(true);
    sh.getRange(head + 1, 3, body.length, 1).setHorizontalAlignment('center');
    let start = 0;
    for (let k = 1; k <= body.length; k++) {
      if (k === body.length || body[k][0] !== body[start][0]) {
        sh.getRange(head + 1 + start, 1, k - start, 1).merge().setWrap(true).setHorizontalAlignment('center').setFontWeight('bold');
        start = k;
      }
    }
    body.forEach((r, k) => { if (r[2] === 'X') sh.getRange(head + 1 + k, 3).setFontColor('#c0352d').setFontWeight('bold'); });
  }
  [96, 380, 80, 240].forEach((w, k) => sh.setColumnWidth(k + 1, w));
  sh.setFrozenRows(head);
}

/** 同一站同一秒送出兩次時，紀錄編號加上 -2、-3…，避免明細混在一起。 */
function uniqueRecordId_(id) {
  const used = {};
  readTable_('檢查紀錄').rows.forEach(r => { used[String(r['紀錄編號'])] = true; });
  let out = id, k = 1;
  while (used[out]) out = id + '-' + (++k);
  return out;
}

function equipItems_() { return items_().filter(i => i.on && i.form === EQUIP_FORM); }

/** 量測值文字：已含單位就不再補單位。 */
function measureText_(v, unit) {
  const m = clean_(v, 30);
  if (!m) return '';
  return '量測值 ' + m + (unit && m.indexOf(unit) < 0 ? ' ' + unit : '');
}

/** 試算表「檢查項目」還沒有營運設備檢查項目時，寫入預設 82 項（只做一次；之後以試算表為準）。 */
function ensureEquipItems_() {
  if (items_().some(i => i.form === EQUIP_FORM)) return 0;
  return withLock_(() => {
    const all = items_();
    if (all.some(i => i.form === EQUIP_FORM)) return 0;
    let order = all.reduce((m, i) => Math.max(m, i.order), 0);
    const rows = [];
    EQUIP_CATS.forEach((c, ci) => c[1].forEach((x, k) => rows.push({
      '項目代碼': equipCode_(ci + 1, k + 1), '表單類型': EQUIP_FORM, '版本': EQUIP_VERSION, '頻率': '不定期', '類別': c[0],
      '檢查內容': x[0], '判定標準': '', '法規依據': '', '風險等級': x[1], '配分': '', '排序': ++order, '啟用': '是'
    })));
    appendObjects_('檢查項目', rows);
    SpreadsheetApp.flush();
    return rows.length;
  });
}

/** 依站點整理營運設備檢查紀錄：最近一次與最近清單。 */
function equipSummary_(codes) {
  const names = {};
  stations_().forEach(s => { names[s.code] = s.name; });
  const recs = readTable_('檢查紀錄').rows
    .filter(r => String(r['表單類型']) === EQUIP_FORM && codes[String(r['站點代碼'])])
    .map(r => ({
      recordId: String(r['紀錄編號']), station: String(r['站點代碼']), stationName: names[String(r['站點代碼'])] || '',
      date: dateStr_(r['檢查日期']), submitted: String(fmtAny_(r['送出時間'])), inspector: String(r['填報人'] || ''),
      ng: Number(r['異常數']) || 0
    }))
    .sort((a, b) => (b.submitted > a.submitted ? 1 : -1));
  const last = {};
  recs.forEach(r => { if (!last[r.station]) last[r.station] = { date: r.date, ng: r.ng, recordId: r.recordId }; });
  return { recent: recs, last: last };
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
  ensureEquipItems_();
  return { stations: stations_(), officers: officers_(), managers: managers_(), items: items_() };
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
  const role = o.role === MANAGER_ROLE ? MANAGER_ROLE : '安保專員';
  return withLock_(() => {
    const cur = (role === MANAGER_ROLE ? managers_() : officers_()).find(x => x.name === (o.originalName || name));
    const row = { '姓名': name, '角色': cur ? cur.role : role, '負責區域': regions.join('、'), 'Email': email };
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
      managers: ['人員', managers_(), x => x.name],
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
  if (!items.some(i => i.form === EQUIP_FORM)) add('提醒', '檢查項目', '尚未建立營運設備檢查項目，第一次開啟「設備檢查」或執行「補上缺少的欄位」時會自動加入 82 項');
  else if (!items.some(i => i.on && i.form === EQUIP_FORM)) add('錯誤', '檢查項目', '營運設備檢查項目全部停用，「設備檢查」無法填寫');

  const s = settings_();
  const pws = ['登入密碼_站長', '登入密碼_安保專員', '登入密碼_系統管理員'].map(k => String(s[k] === undefined ? '' : s[k]).trim());
  if (pws.some(p => !p)) add('錯誤', '設定', '登入密碼有空白');
  if (new Set(pws).size !== pws.length) add('錯誤', '設定', '三組登入密碼不能相同');
  const mpw = String(s['登入密碼_油品主管'] || s['登入密碼_油品主管經理'] || '').trim();
  if (mpw) {
    if (pws.indexOf(mpw) >= 0) add('錯誤', '設定', '油品主管的登入密碼不能和其他角色相同');
    if (!managers_().some(m => m.on)) add('提醒', '人員', '已設定油品主管登入密碼，但還沒有啟用中的油品主管');
  } else if (managers_().some(m => m.on)) add('提醒', '設定', '已建立油品主管，但「設定」工作表沒有「登入密碼_油品主管」，油品主管無法登入');
  if (pws.indexOf('12345') >= 0 || pws.indexOf('98765') >= 0 || pws.indexOf('77777') >= 0) add('提醒', '設定', '仍在使用初始密碼，正式上線前請修改');
  if (!getProp_('PHOTO_FOLDER_ID')) add('提醒', '指令碼屬性', '尚未設定 PHOTO_FOLDER_ID，第一次上傳照片時會在試算表所在資料夾自動建立「巡檢照片」');
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
  const n = ensureEquipItems_();
  if (n) added.push('檢查項目：加入營運設備檢查 ' + n + ' 項');
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

function officers_() { return people_(['安保專員']); }
function managers_() { return people_(MANAGER_ROLE_NAMES); }

function people_(roles) {
  return readTable_('人員').rows
    .filter(r => roles.indexOf(String(r['角色'] || '').trim()) >= 0)
    .map(r => ({
      role: String(r['角色'] || '').trim(),
      name: String(r['姓名'] || '').trim(),
      regions: /全區|全部/.test(String(r['負責區域'] || '')) ? REGIONS.slice() : String(r['負責區域'] || '').split(/[、,，\s]+/).filter(x => REGIONS.indexOf(x) >= 0),
      email: String(r['Email'] || '').trim(), on: yes_(r['啟用']), _row: r._row
    })).filter(o => o.name);
}

function items_() {
  return readTable_('檢查項目').rows.map(r => ({
    code: String(r['項目代碼'] || '').trim(), form: String(r['表單類型'] || '').trim(),
    version: String(r['版本'] || ''), freq: String(r['頻率'] || '').trim(), cat: String(r['類別'] || '').trim(),
    text: String(r['檢查內容'] || '').trim(), std: String(r['判定標準'] || ''), law: String(r['法規依據'] || ''),
    risk: String(r['風險等級'] || '').trim(), pts: r['配分'] === '' || r['配分'] === null ? null : Number(r['配分']),
    order: Number(r['排序']) || 0, on: yes_(r['啟用']), suggest: String(r['改善建議'] || '').trim(), _row: r._row
  })).filter(i => i.code).sort((a, b) => a.order - b.order);
}

function scopeStations_(sess) {
  const all = activeStations_();
  if (sess.role === '站長') return all.filter(s => s.code === sess.station);
  if (sess.role === '安保專員' || sess.role === MANAGER_ROLE) return all.filter(s => sess.regions.indexOf(s.region) >= 0);
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
