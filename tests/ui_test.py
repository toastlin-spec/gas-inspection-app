import sys, struct, zlib
from playwright.sync_api import sync_playwright

OUT = './'
def png(path):
    # 產生一張 40x30 的測試照片
    w, h = 40, 30
    raw = b''.join(b'\x00' + bytes([200, 60, 50] * w) for _ in range(h))
    def chunk(t, d): return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    open(path, 'wb').write(b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b''))
png(OUT + 'photo.png')

errors, steps = [], []
def ok(m): steps.append(m); print('  ✓', m)

with sync_playwright() as p:
    b = p.chromium.launch()
    pg = b.new_page(viewport={'width': 390, 'height': 844}, device_scale_factor=2)
    pg.on('pageerror', lambda e: errors.append(str(e)))
    pg.on('console', lambda m: errors.append(m.text) if m.type == 'error' and 'fonts.g' not in m.text else None)
    pg.goto('http://localhost:8765/')
    pg.wait_for_selector('#pw')

    pg.fill('#pw', '11111'); pg.click('text=下一步'); pg.wait_for_selector('.msg')
    ok('錯誤密碼顯示：' + pg.inner_text('.msg'))
    pg.fill('#pw', '12345'); pg.click('text=下一步'); pg.wait_for_selector('#lgSt')
    pg.select_option('#lgSt', 'C04')
    assert pg.input_value('#lgName') == '李丞芯'; ok('選台中站，姓名自動帶入李丞芯')
    pg.screenshot(path=OUT + '1_login.png')
    pg.click('text=進入'); pg.wait_for_selector('.todo')
    ok('站長首頁：' + pg.inner_text('main h1'))
    pg.screenshot(path=OUT + '2_home.png')

    pg.click('button.todo[data-freq="每日"]'); pg.wait_for_selector('#it-SC-D01')
    pg.click('#it-SC-D01 button.ng'); pg.fill('#n-SC-D01', '加油島地面油漬未清除')
    pg.set_input_files('#p-SC-D01', OUT + 'photo.png'); pg.wait_for_selector('#it-SC-D01 img.thumb')
    ok('選異常、填說明、拍照後顯示縮圖')
    pg.click('text=送出每日檢查'); pg.wait_for_selector('.toast:not([hidden])')
    ok('未完成就送出提示：' + pg.inner_text('#toast'))
    pg.click('#it-SC-D02 button.na')
    pg.screenshot(path=OUT + '3_check.png', full_page=True)
    pg.click('text=送出每日檢查'); pg.wait_for_selector('.done-card')
    ok('送出成功：' + pg.inner_text('.done-card h1') + '｜' + pg.inner_text('.done-card p >> nth=1'))
    pg.screenshot(path=OUT + '4_done.png')

    pg.click('nav a[href="#issues"]'); pg.wait_for_selector('.issue')
    pg.click('.issue'); pg.wait_for_selector('.sheet'); ok('異常清單與詳情：' + pg.inner_text('.sheet h1'))
    pg.screenshot(path=OUT + '5_issue.png')
    pg.click('.iconbtn')

    pg.click('nav a[href="#home"]'); pg.wait_for_selector('.todo.done')
    ok('回首頁顯示今日已送出')

    pg.click('text=登出'); pg.wait_for_selector('#pw')
    pg.fill('#pw', '98765'); pg.click('text=下一步'); pg.wait_for_selector('.who')
    pg.click('.who button:has-text("林佳賢")'); pg.click('text=進入'); pg.wait_for_selector('.grp-title')
    ok('安保專員首頁：' + pg.inner_text('main h2 >> nth=0'))
    pg.screenshot(path=OUT + '6_officer.png')

    pg.click('text=登出'); pg.wait_for_selector('#pw')
    pg.fill('#pw', '77777'); pg.click('text=下一步'); pg.wait_for_selector('nav a[href="#manage"]')
    pg.click('nav a[href="#manage"]'); pg.wait_for_selector('.list-row')
    pg.click('text=新增站點'); pg.wait_for_selector('#fsName')
    pg.select_option('#fsRegion', '南區'); pg.fill('#fsGroup', '南一組'); pg.fill('#fsName', '測試站'); pg.fill('#fsMgr', '王測試')
    pg.click('.sheet button.primary'); pg.wait_for_selector('text=已儲存站點')
    pg.click('button.chip[data-r="南區"]'); pg.wait_for_selector('text=S28')
    ok('管理員新增南區站點 S28')
    pg.screenshot(path=OUT + '7_manage.png')
    pg.click('button[data-t="資料檢查"]'); pg.click('text=開始檢查'); pg.wait_for_selector('.risk')
    ok('資料檢查列出 %d 項提醒' % pg.locator('section .list-row').count())
    pg.screenshot(path=OUT + '8_datacheck.png')
    b.close()

print('頁面錯誤：', errors if errors else '無')
sys.exit(1 if errors else 0)
