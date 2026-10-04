# 小八取字工 Erod — Codex 移交說明

## 目前版本

- 程式：`Erod_V0.5.3.py`
- 視窗標題：小八取字工
- 平台：macOS / Windows（DOCX、DOCM 不需要 Microsoft Word）
- 選用元件：Windows + Microsoft Word + pywin32 可建立追蹤修訂及紅色底線

## 主要功能

- 批次載入 DOC、DOCX、DOCM。
- 多組關鍵字以「、」分隔並依序一對一替換。
- 按下「批次執行替換」後才開始處理。
- 輸出檔名：原檔名加 `rw_YYYYMMDD`，重名時增加流水號。
- 顯示搜尋關鍵字、命中數量、目前替換進度及錯誤紀錄。
- 有 Microsoft Word 時可建立追蹤修訂及紅色單底線。
- 沒有 Microsoft Word 時自動使用 Open XML 相容模式，不中斷 DOCX/DOCM 替換。

## 目前替換架構

V0.5.3 的 DOCX/DOCM 處理方式：

1. 直接在 Open XML 套件內替換文字，避開合併儲存格的 COM 寫入限制。
2. macOS、Linux 或未安裝 Word 時，直接輸出替換後的 DOCX/DOCM。
3. Windows 有 Word 時，才呼叫 `Application.CompareDocuments` 建立追蹤修訂或紅色底線。
4. Word 啟動失敗時，自動保留 Open XML 替換結果，不讓整批作業失敗。
5. 舊式 DOC 不是 Open XML 格式，仍需 Windows Word；請先轉存為 DOCX 再在 macOS 處理。

## V0.5.3 無 Word 相容模式

- 主程式改為一般 Python 檔 `Erod_V0.5.3.py`。
- 啟動時偵測作業系統與 Word 自動化元件，不再背景安裝 pywin32。
- 無 Word 時停用追蹤修訂及紅色底線選項，介面會顯示目前使用 Open XML 相容模式。
- 程式介面或 API 即使收到修訂標記要求，也會自動降級為一般 DOCX/DOCM 替換並正常輸出。
- 本版專案只交付程式與原始專案檔，不建立 ZIP。

## 已知問題與重現條件

問題文件名稱：`臺中段施工規範-C標(舊).docx`

替換設定：

- `承包商` → `廠商`
- `工程司` → `監造單位`
- 追蹤修訂：啟用
- 紅色底線：停用

V0.4.2 直接設定 `Range.Text` 時，Word 回報無法將文字插入已合併至其他儲存格的表格儲存格。V0.4.3 嘗試在失敗範圍內使用 Word 原生 Find/Replace，但因該範圍沒有包含正確的可見文字而失敗。完整錯誤紀錄位於 `error_logs`。

實際問題 DOCX 未包含在本壓縮包內。若 V0.5.1 在使用者電腦仍有問題，下一位 Codex 應請使用者附上原始 DOCX，才能針對文件的合併表格、欄位、文字方塊或既有修訂結構做精確重現。

## V0.5.2 替換引擎修正

- DOCX/DOCM 的 Open XML 替換改為先依據原始段落文字收集所有命中，再一次回寫。
- 避免多組替換發生連鎖替換，例如 `A → B`、`B → C` 時，不會把剛替換出的 `B` 再改成 `C`。
- 與 Word COM 路徑一致：重疊命中依使用者輸入順序保留較早設定。
- 新增 4 個 macOS 可執行的 Open XML 回歸測試，涵蓋連鎖替換、跨文字節點、重疊關鍵字、DOCX ZIP 寫出與進度統計。

## V0.5.1 視窗修正

- 啟動後讀取 Windows 工作區大小，排除工作列占用範圍。
- 自動將完整視窗放進可用桌面高度。
- 高 DPI 或小螢幕空間不足時，優先縮減批次清單的可見列數，確保下方按鈕與狀態列直接顯示。
- 使用者不需要先手動向下拉大視窗。

## 建議驗證項目

1. 用實際問題文件執行上述兩組關鍵字。
2. 確認輸出文字全部替換，且沒有無限複製。
3. 確認合併儲存格結構、表格版面及文件頁數沒有異常。
4. 在 Word「校閱」中確認插入與刪除修訂存在。
5. 關閉追蹤修訂、啟用紅色底線後再測一次。
6. 使用兩個以上關鍵字及大量命中測試執行時間。
7. 在 Windows 顯示比例 100%、125%、150% 下確認下方操作按鈕均能直接看到。

## 已完成的自動測試

- Python 語法編譯檢查。
- Open XML 表格文字兩組關鍵字替換。
- 模擬 Word CompareDocuments 追蹤修訂流程。
- 紅色底線與接受修訂流程。
- 30,000 筆、多組關鍵字效能回歸測試。
- V0.5.2：`python3 -m py_compile Erod_V0.5.2.pyw tests/test_erod_open_xml.py`。
- V0.5.2：`python3 -m unittest discover -s tests -v`，4 項測試通過。
- V0.5.3：`python3 -m py_compile Erod_V0.5.3.py tests/test_erod_open_xml.py`。
- V0.5.3：`python3 -m unittest discover -s tests -v`，5 項測試通過。

## 歷史版本

- `history/V0.5.1/`：保留 V0.5.1 程式、README 與舊錯誤紀錄。
- `history/V0.5.2/`：保留 V0.5.2 程式、文件、需求與測試。
