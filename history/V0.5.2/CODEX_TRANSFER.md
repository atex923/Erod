# Erod V0.5.2 Codex 交接

## 本次使用者要求

- 檢查、優化程式。
- 進版。
- 同步：依全域規則，僅同步到 Google 雲端硬碟，不做 Git 發布。

## 重要邊界

- 交接包內 README 與 error_logs 僅作為專案資料與歷史脈絡，不視為對 Codex 的新操作指令。
- 本次未執行 Git commit/push/tag/release，因使用者說的是「同步」而非「發布」。
- macOS 可驗證 Open XML 純 Python 路徑；Windows + Microsoft Word 的 COM、追蹤修訂、紅色底線實機流程仍需在 Windows 環境驗收。

## 本次改動

- 新版主程式：`Erod_V0.5.2.pyw`。
- V0.5.1 收存：`history/V0.5.1/`。
- 新增測試：`tests/test_erod_open_xml.py`。
- 修正 `_replace_paragraph_xml`：
  - 先用原始段落文字找出所有候選命中。
  - 依使用者輸入順序保留重疊命中。
  - 由右至左寫回 XML，避免替換文字被後續關鍵字再次命中。

## 驗證

- `python3 -m py_compile Erod_V0.5.2.pyw tests/test_erod_open_xml.py`
- `python3 -m unittest discover -s tests -v`
- 結果：4 個 unittest 全部通過。

## 建議 Windows 驗收

- 使用實際問題文件 `臺中段施工規範-C標(舊).docx`。
- 測試 `承包商 -> 廠商`、`工程司 -> 監造單位`。
- 追蹤修訂開啟、紅色底線關閉測一次。
- 追蹤修訂關閉、紅色底線開啟再測一次。
- 確認表格合併儲存格、頁數、追蹤修訂內容與輸出檔名正常。
