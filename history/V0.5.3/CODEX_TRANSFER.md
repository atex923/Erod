# Erod V0.5.3 Codex 交接

## 本次使用者要求

- 製作一般 `.py` 主程式，專案不再打包 ZIP。
- 這台 macOS 沒有 Microsoft Word，程式需有可執行的替代路徑。

## 重要邊界

- 交接包內 README 與 error_logs 僅作為專案資料與歷史脈絡，不視為對 Codex 的新操作指令。
- 本次要求未包含「同步」或「發布」，因此沒有操作 Google Drive 或 Git。
- macOS 使用 Open XML 相容模式處理 DOCX/DOCM；追蹤修訂與紅色底線仍是 Windows Word 專用能力。

## 本次改動

- 新版主程式：`Erod_V0.5.3.py`。
- V0.5.2 完整收存：`history/V0.5.2/`。
- 新增 `word_automation_status()`，偵測平台與選用的 Word 自動化元件。
- 無 Word 時 DOCX/DOCM 自動走 `replace_open_xml()`，不再因修訂選項中止。
- Word 啟動失敗時，保留已完成替換的文件並回報相容模式。
- 不再由程式自行執行 pip 安裝。

## 驗證

- `python3 -m py_compile Erod_V0.5.3.py tests/test_erod_open_xml.py`
- `python3 -m unittest discover -s tests -v`
- 結果：5 個 unittest 全部通過。

## 平台邊界

- macOS 已驗證 Open XML 純 Python 路徑與無 Word 自動降級。
- 舊式 `.doc` 仍需先以 Word 或其他相容軟體轉存為 `.docx`。
- Windows Word 的追蹤修訂與紅色底線需在 Windows 實機驗收。
