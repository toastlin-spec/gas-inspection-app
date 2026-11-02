# 測試

用 Node 模擬 Apps Script 的 SpreadsheetApp、DriveApp、CacheService 等服務，載入資料範本測試 `src/Code.gs`。

1. 把 `data/加油站巡檢App_資料範本.xlsx`（或含正式站點資料的版本）轉成 `tests/workbook.json`，每個工作表一個二維陣列。此檔含人員資料，不放進儲存庫。
2. 後端情境測試：`node tests/test.js`
3. 手機畫面測試：先執行 `node tests/server.js`，再執行 `python3 tests/ui_test.py`（需要 Playwright）。
