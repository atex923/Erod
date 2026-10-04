# 小八取字工 Erod

Erod 是一套以 Python 製作的 Word 批次文字替換工具，目前版本為 **V0.5.4**。它可批次處理 DOCX／DOCM，支援多組關鍵字、替換文字自選顏色，以及無 Microsoft Word 的 Open XML 相容模式。

## 最新版

- [Erod_V0.5.4.py](Erod_V0.5.4.py)：macOS、Windows 與命令列啟動入口。
- [Erod_V0.5.4.pyw](Erod_V0.5.4.pyw)：Windows 雙擊啟動入口，不顯示命令列視窗。
- [版本紀錄](RELEASE_NOTES.md)
- [詳細技術說明](README_Codex.md)

`.py` 與 `.pyw` 內容完全一致。

## 程式特色

- 批次載入 DOC、DOCX、DOCM 文件。
- 使用「、」、分號或換行輸入多組關鍵字與替換文字。
- DOCX／DOCM 直接以 Open XML 替換，不需要安裝 Microsoft Word。
- 避免多組規則發生連鎖替換，並依輸入順序處理重疊關鍵字。
- 支援跨文字節點與合併儲存格內的文字替換。
- 可只替換新文字套用自選顏色，預設為紅色。
- Windows 安裝 Word 與 pywin32 時，可建立追蹤修訂及紅色底線。
- Word 無法使用時自動保留一般替換結果，不中斷整批作業。

## 安裝與執行

需要 Python 3.10 以上版本。

```bash
python3 -m pip install -r requirements.txt
python3 Erod_V0.5.4.py
```

Windows 可直接雙擊 `Erod_V0.5.4.pyw`。`tkinterdnd2` 用於拖放檔案；未安裝時仍可用按鈕選擇檔案。`pywin32` 僅在 Windows 的 Microsoft Word 追蹤修訂功能中使用。

舊式 `.doc` 不是 Open XML 格式，沒有 Windows Word 時請先轉存為 `.docx`。

## 輸出方式

輸出文件會放在來源文件旁，檔名格式為：

```text
原檔名rw_YYYYMMDD.docx
```

若同名檔案已存在，程式會自動增加流水號，不覆蓋原始文件。

## 檔案與版本歷史

- 根目錄保留 V0.5.4 最新程式與文件。
- [history/](history/) 保存 V0.5.1 至 V0.5.3 的完整版本資料。
- [tests/](tests/) 保存 Open XML 回歸測試。

## 驗證

V0.5.4 已通過 Python 語法檢查及 8 項單元測試：

```bash
python3 -B -m unittest discover -s tests -v
```

macOS 已驗證無 Word 的 Open XML 路徑。Windows Word 的 COM、追蹤修訂與紅色底線仍需在 Windows 實機驗證。
