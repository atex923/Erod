# 小八取字工 Erod 版本紀錄

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
