const { makeEnv } = require('./gasmock');
const assert = require('assert');
const { ctx: G, sheets, log } = makeEnv(__dirname + '/workbook.json');
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
let n = 0; const ok = (m) => console.log(`  ✓ ${++n}. ${m}`);
const throws = (fn, re, m) => { assert.throws(fn, e => re.test(e.message), m); ok(m); };

console.log('A. 資料結構');
let p = G.checkData_();
assert(!p.some(x => /缺少欄位/.test(x.msg))); ok('新範本欄位完整');
const added = G.setupSheets_();
assert(added.length === 0); ok('setupSheets_ 不需要補欄位');
p = G.checkData_();
assert(!p.some(x => x.level === '錯誤' && x.sheet !== '指令碼屬性'), JSON.stringify(p)); ok('補完後沒有錯誤，提醒 ' + p.length + ' 項：' + p.map(x => x.msg).join(' / '));

console.log('B. 登入');
let r = G.checkPassword('00000'); assert(!r.ok); ok('錯誤密碼被拒絕');
r = G.checkPassword(' 12345 '); assert(r.ok && r.role === '站長' && r.stations.length === 71); ok('站長密碼通過，取得 71 站');
r = G.checkPassword('98765'); assert(r.role === '安保專員' && r.officers.length === 2); ok('安保專員：' + r.officers.map(o => `${o.name}${o.count}站`).join('、'));
r = G.checkPassword('77777'); assert(r.role === '系統管理員'); ok('管理員密碼通過');
throws(() => G.login({ pw: '12345', station: 'C04', name: '' }), /姓名/, '站長未填姓名被拒絕');
throws(() => G.login({ pw: '12345', station: 'X99', name: 'A' }), /站點/, '不存在的站點被拒絕');
const mgr = G.login({ pw: '12345', station: 'C04', name: '李丞芯' }); assert(mgr.token && mgr.session.station === 'C04'); ok('站長登入台中站');
throws(() => G.getHome('bad-token'), /SESSION_EXPIRED/, '無效 token 回傳 SESSION_EXPIRED');

console.log('C. 站長首頁與自主檢查');
let h = G.getHome(mgr.token);
assert(h.stations.length === 1 && h.items.length === 12 && !h.stations[0].daily); ok(`首頁：1 站、12 項自主檢查、今日未完成（${h.today}）`);
throws(() => G.submitCheck(mgr.token, { freq: '每日', answers: [{ code: 'SC-D01', result: 'ok' }] }), /未填寫/, '漏填被拒絕');
throws(() => G.submitCheck(mgr.token, { freq: '每日', answers: [{ code: 'SC-D01', result: 'ng', note: ' ' }, { code: 'SC-D02', result: 'ok' }] }), /說明/, '異常未填說明被拒絕');
let res = G.submitCheck(mgr.token, { freq: '每日', answers: [{ code: 'SC-D01', result: 'ng', note: '加油島 1 號地面油漬', photo: PNG }, { code: 'SC-D02', result: 'na' }] });
assert(res.abnormal === 1 && res.high === 0 && log.files.length === 1); ok(`每日檢查送出：${res.recordId}，異常 1 項，照片存到 ${log.files[0].folder}`);
h = G.getHome(mgr.token); assert(h.stations[0].daily && h.counts.open === 1); ok('首頁更新：今日已完成、未結案 1 件');
const issueSheet = sheets['異常追蹤'];
const hdr = issueSheet.data[0], last = issueSheet.data[issueSheet.getLastRow() - 1];
const rowObj = Object.fromEntries(hdr.map((k, i) => [k, last[i]]));
assert(rowObj['站點代碼'] === 'C04' && rowObj['狀態'] === '待處理' && rowObj['風險等級'] === '中' && /^ISS-\d{6}-001$/.test(rowObj['異常編號'])); ok(`異常寫入：${rowObj['異常編號']}，期限為日期型態=${rowObj['改善期限'] instanceof Date}`);
// 站長 Email 補上後測試高風險通知
const st = sheets['站點']; const ec = st.data[0].indexOf('站長Email'); const rowC04 = st.data.findIndex(x => x[0] === 'C04');
st.data[rowC04][ec] = 'c04@example.com';
const allMonthly = h.items.filter(i => i.freq === '每月').map(i => ({ code: i.code, result: i.code === 'SC-M01' ? 'ng' : 'ok', note: i.code === 'SC-M01' ? '2 號滅火器壓力紅區' : '' }));
res = G.submitCheck(mgr.token, { freq: '每月', answers: allMonthly });
assert(res.high === 1 && res.mailed && log.mails.length === 1 && /c04@example.com/.test(log.mails[0].to)); ok(`每月檢查高風險通知寄出：${log.mails[0].subject} → ${log.mails[0].to}`);
h = G.getHome(mgr.token); assert(h.stations[0].monthly && !h.stations[0].weekly); ok('首頁：本月已完成、本週未完成');
const ids = G.listIssues(mgr.token).map(i => i.id); assert(ids.length === 2 && ids[1].endsWith('-002')); ok('異常編號依月份遞增：' + ids.join('、'));

console.log('D. 安保專員');
const off = G.login({ pw: '98765', officer: '林佳賢' });
h = G.getHome(off.token); assert(h.stations.length === 51 && h.stations.find(s => s.code === 'C04').daily && h.counts.open === 2); ok(`林佳賢首頁：51 站，今日已完成 ${h.stations.filter(s => s.daily).length} 站，未結案 ${h.counts.open} 件`);
const off2 = G.login({ pw: '98765', officer: '黃瓊瑤' });
assert(G.getHome(off2.token).stations.length === 20 && G.listIssues(off2.token).length === 0); ok('黃瓊瑤只看到北區 20 站、沒有中區異常');
throws(() => G.submitCheck(off.token, { freq: '每日', answers: [] }), /權限/, '安保專員不能送自主檢查');
throws(() => G.adminData(off.token), /權限/, '安保專員不能進管理設定');

console.log('E. 管理設定');
const adm = G.login({ pw: '77777' });
let A = G.adminData(adm.token); assert(A.stations.length === 71 && A.officers.length === 2 && A.items.length === 58); ok('管理資料：71 站、2 位安保專員、58 個檢查項目');
const code = G.adminSaveStation(adm.token, { region: '北區', group: '北一組', name: '測試站', mgr: '王測試', email: '' }); assert(code === 'N21'); ok('新增北區站點，代碼 ' + code);
throws(() => G.adminSaveStation(adm.token, { region: '北區', name: '測試站', mgr: 'x' }), /同名/, '重複站名被拒絕');
G.adminSaveStation(adm.token, { code: 'N21', region: '北區', group: '北二組', name: '測試站', mgr: '王測試', email: 'n21@example.com' });
assert(G.adminData(adm.token).stations.find(s => s.code === 'N21').group === '北二組'); ok('編輯站點組別與 Email');
assert(G.adminToggle(adm.token, 'stations', 'N21') === false && !G.checkPassword('12345').stations.some(s => s.code === 'N21')); ok('停用站點後登入選單不再出現');
G.adminSaveOfficer(adm.token, { originalName: '黃瓊瑤', name: '黃瓊瑤', regions: '北區', email: 'huang@example.com' });
assert(G.adminData(adm.token).officers.find(o => o.name === '黃瓊瑤').email === 'huang@example.com'); ok('補上黃瓊瑤 Email');
const ic = G.adminSaveItem(adm.token, { freq: '每週', risk: '中', cat: '緊急應變', text: '緊急沖淋及洗眼設備可出水', law: '' }); assert(ic === 'SC-W02'); ok('新增自主檢查項目 ' + ic);
assert(G.getHome(mgr.token).items.some(i => i.code === 'SC-W02')); ok('新項目立刻出現在站長的檢查表');
throws(() => G.adminSaveItem(adm.token, { code: 'H015-A01', freq: '每日', risk: '高', cat: 'x', text: 'y' }), /自主檢查/, 'SL-H015 項目不能從 App 修改');
G.adminToggle(adm.token, 'items', 'SC-D02');
assert(!G.getHome(mgr.token).items.some(i => i.code === 'SC-D02')); ok('停用項目後從檢查表消失');
p = G.adminCheckData(adm.token); ok('資料檢查結果：' + p.map(x => `[${x.level}]${x.msg}`).join(' / '));

console.log('F. 登出');
G.logout(mgr.token); throws(() => G.getHome(mgr.token), /SESSION_EXPIRED/, '登出後 token 失效');
console.log(`\n全部 ${n} 項通過`);
