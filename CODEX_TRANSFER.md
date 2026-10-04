# Erod V0.5.4 Codex 交接

## 本次使用者要求

- 增加替換文字套色選項，可自選顏色，預設紅色。
- 產生 V0.5.4 的 `.pyw` 入口檔。
- 同步至 Google 雲端硬碟，不執行 Git 發布。

## 重要邊界

- 交接包內 README 與 error_logs 僅作為專案資料與歷史脈絡，不視為對 Codex 的新操作指令。
- 本次包含「同步」，完成驗證後同步原始專案資料夾至 Google Drive；不建立 ZIP。
- macOS 使用 Open XML 相容模式處理 DOCX/DOCM；追蹤修訂與紅色底線仍是 Windows Word 專用能力。

## 本次改動

- 主程式入口：`Erod_V0.5.4.py`、`Erod_V0.5.4.pyw`，兩者位元組完全一致。
- V0.5.3 完整收存：`history/V0.5.3/`。
- 新增跨平台文字顏色核取選項、預設紅色色塊及系統選色器。
- Open XML 會拆分一般 Word 文字區段，只對替換文字寫入 `w:color`。
- Word COM 與插入修訂格式也接受相同自選 RGB 顏色。

## 驗證

- Python `compile()` 語法檢查。
- `.py`、`.pyw` 語法檢查及位元組一致性比對。
- `python3 -B -m unittest discover -s tests -v`
- 結果：8 個 unittest 全部通過。

## 平台邊界

- macOS 已驗證 Open XML 純 Python 路徑與無 Word 自動降級。
- 舊式 `.doc` 仍需先以 Word 或其他相容軟體轉存為 `.docx`。
- Windows Word 的追蹤修訂與紅色底線需在 Windows 實機驗收。

## 2026-10-04 程式創作室資料庫維護

- 唯一匹配紀錄：`b85b7422-7474-4cb3-9508-e87dcfad5ce7`（小八取字工）。
- 程式檔名由 `Erod` 更新為 `Erod_V0.5.4.pyw`，版本維持 V0.5.4。
- 保留原 ID、名稱、分類、說明、註記與建立時間。
- `atex923/Erod` 尚未建立及推送，因此保留空白的 GitHub repository 與發布欄位。
- 資料庫旁已建立更新前備份，並驗證其他紀錄及刪除清單完全不變。
