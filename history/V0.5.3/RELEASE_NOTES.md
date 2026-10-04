# 小八取字工 Erod 版本紀錄

## V0.5.3 - 2026-08-24

- 主程式由 `Erod_V0.5.2.pyw` 進版為 `Erod_V0.5.3.py`。
- DOCX/DOCM 在 macOS、Linux、未安裝 Word 或缺少 pywin32 時，自動使用 Open XML 相容模式完成替換。
- 無 Word 時不再卡在追蹤修訂功能；介面會停用 Word 專用選項並顯示替代模式。
- Windows 有 pywin32 但 Word 無法啟動時，保留已完成的 Open XML 替換結果。
- 移除程式自行背景安裝 pywin32 的行為。
- 新增無 Word 降級輸出測試；5 項 unittest 全部通過。
- 本版依使用者要求不建立 ZIP。

## V0.5.2 - 2026-08-24

- 進版：`Erod_V0.5.1.pyw` -> `Erod_V0.5.2.pyw`。
- 修正 DOCX/DOCM Open XML 多組替換的連鎖替換風險。
- Open XML 替換現在以原始段落文字建立候選命中，依使用者輸入順序處理重疊關鍵字，再由右至左寫回 XML 文字節點。
- 更新合併儲存格錯誤建議，避免仍描述舊版 Word 原生替換流程。
- 新增 `tests/test_erod_open_xml.py`，涵蓋連鎖替換、跨節點替換、重疊關鍵字優先序、DOCX ZIP 寫出與搜尋統計。
- 驗證通過：
  - `python3 -m py_compile Erod_V0.5.2.pyw tests/test_erod_open_xml.py`
  - `python3 -m unittest discover -s tests -v`

## V0.5.1 - 2026-08-24

- 修正啟動視窗高度，依 Windows 工作區排除工作列後自動縮放。
- 高 DPI 或小螢幕時縮減批次清單列數，讓底部按鈕與狀態列可直接顯示。
- 舊版檔案已收存於 `history/V0.5.1/`。
