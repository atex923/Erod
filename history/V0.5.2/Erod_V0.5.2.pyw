from __future__ import annotations

import bisect
import html
import importlib
import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import traceback
import zipfile
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


APP_NAME = "小八取字工"
VERSION = "V0.5.2"
SUPPORTED_EXTENSIONS = {".doc", ".docx", ".docm"}
MAX_REPLACEMENTS_PER_DOCUMENT = 100_000


def _emit_progress(progress_callback, stage: str, data: object) -> None:
    if progress_callback is None:
        return
    try:
        progress_callback(stage, data)
    except Exception:
        # Status reporting must never interrupt document processing.
        pass


try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    DND_AVAILABLE = True
except ImportError:
    DND_FILES = None
    TkinterDnD = None
    DND_AVAILABLE = False


def split_items(value: str) -> list[str]:
    """Accept the Chinese separator, line breaks, and semicolons."""
    return [item.strip() for item in re.split(r"[、；;\r\n]+", value) if item.strip()]


def make_replacements(find_text: str, replace_text: str) -> list[tuple[str, str]]:
    keywords = split_items(find_text)
    replacements = split_items(replace_text)
    if not keywords:
        raise ValueError("請先輸入要尋找的關鍵字。")
    if not replacements:
        # An intentionally blank replacement is represented by a single space.
        if replace_text == "":
            replacements = [""]
        else:
            raise ValueError("請輸入替換文字。")
    if len(replacements) == 1:
        return [(keyword, replacements[0]) for keyword in keywords]
    if len(keywords) != len(replacements):
        raise ValueError(
            f"尋找關鍵字有 {len(keywords)} 組，替換文字有 {len(replacements)} 組，"
            "兩邊數量必須相同；或替換文字只輸入 1 組。"
        )
    return list(zip(keywords, replacements))


def explain_error(error_text: str) -> tuple[str, str]:
    """Turn common Word/Python failures into short Traditional Chinese guidance."""
    lowered = error_text.lower()
    if "沒有找到指定的關鍵字" in error_text:
        return (
            "文件內找不到關鍵字",
            "請確認文字、空格及全形／半形符號完全相同，也要確認目前使用的不是掃描圖片。",
        )
    if "permission denied" in lowered or "拒絕存取" in error_text:
        return (
            "檔案沒有寫入權限",
            "請關閉正在開啟此檔案的 Word 視窗，並確認資料夾不是唯讀或受權限保護。",
        )
    if "badzipfile" in lowered or "not a zip file" in lowered:
        return (
            "Word 檔案格式損壞或副檔名不正確",
            "請先用 Word 開啟原檔並另存為新的 DOCX，再重新處理。",
        )
    if "無法啟動 microsoft word" in lowered:
        return (
            "無法啟動 Microsoft Word",
            "請確認已安裝 Windows 桌面版 Word，並先手動開啟一次 Word 後再重試。",
        )
    if "pywin32" in lowered or "pythoncom" in lowered or "win32com" in lowered:
        return (
            "缺少 Word 自動化元件",
            "請在目前使用的 Python 執行：pip install pywin32，完成後重新開啟程式。",
        )
    if "password" in lowered or "protected" in lowered or "保護" in error_text:
        return (
            "文件受到密碼或編輯保護",
            "請先在 Word 中解除文件保護，再把解除保護後的檔案加入清單。",
        )
    if "rpc" in lowered or "call was rejected" in lowered or "被呼叫端拒絕" in error_text:
        return (
            "Microsoft Word 忙碌或沒有回應",
            "請關閉 Word 的對話框及其他 Word 視窗，必要時結束 WINWORD.EXE 後再試。",
        )
    if "100,000" in error_text:
        return (
            "關鍵字命中數量超過安全限制",
            "請縮小關鍵字範圍或分次處理，以免替換到非預期的大量文字。",
        )
    if (
        "已合併至其他儲存格" in error_text
        or "merged into other cells" in lowered
        or "merged cell" in lowered
    ):
        return (
            "合併儲存格內的文字無法替換",
            "新版 DOCX/DOCM 會優先使用 Open XML 避開儲存格限制；若仍失敗，請檢查該表格是否損壞或另存新檔後重試。",
        )
    if "no such file" in lowered or "找不到" in error_text:
        return (
            "來源檔案不存在或已被移動",
            "請從批次清單移除該項目，重新加入目前實際存在的 Word 檔案。",
        )
    first_line = next((line.strip() for line in error_text.splitlines() if line.strip()), "未知錯誤")
    return (
        f"未分類錯誤：{first_line[:80]}",
        "請依錯誤紀錄中的技術細節檢查；也可先用 Word 另存新檔後再重試。",
    )


def write_error_log(
    files: list[Path],
    replacements: list[tuple[str, str]],
    track_changes: bool,
    red_underline: bool,
    completed: list[tuple[Path, Path, int]],
    failed: list[tuple[Path, str, str, str]],
) -> Path | None:
    if not failed:
        return None
    now = datetime.now()
    lines = [
        f"{APP_NAME} 錯誤輸出說明",
        f"程式版本：{VERSION}",
        f"處理時間：{now:%Y-%m-%d %H:%M:%S}",
        f"總檔案數：{len(files)}",
        f"成功：{len(completed)}　失敗：{len(failed)}",
        f"追蹤修訂：{'啟用' if track_changes else '停用'}",
        f"紅色底線：{'啟用' if red_underline else '停用'}",
        "",
        "替換設定：",
    ]
    for find_value, replace_value in replacements:
        lines.append(f"  「{find_value}」 → 「{replace_value}」")
    lines.extend(["", "失敗檔案："])
    for index, (path, summary, suggestion, technical) in enumerate(failed, start=1):
        lines.extend(
            [
                "",
                f"[{index}] {path}",
                f"錯誤說明：{summary}",
                f"處理建議：{suggestion}",
                "技術細節：",
                technical.rstrip(),
            ]
        )
    if completed:
        lines.extend(["", "成功輸出："])
        for source, output, count in completed:
            lines.append(f"  {source.name} → {output.name}（替換 {count} 處）")

    filename = f"Erod_error_{now:%Y%m%d_%H%M%S}.txt"
    directories: list[Path] = []
    if files:
        directories.append(files[0].parent)
    directories.extend([Path.cwd(), Path.home() / "Desktop"])
    seen: set[str] = set()
    for directory in directories:
        key = os.path.normcase(str(directory))
        if key in seen:
            continue
        seen.add(key)
        try:
            directory.mkdir(parents=True, exist_ok=True)
            log_path = directory / filename
            base_stem = log_path.stem
            suffix = 2
            while log_path.exists():
                log_path = directory / f"{base_stem}_{suffix:02d}.txt"
                suffix += 1
            log_path.write_text("\n".join(lines), encoding="utf-8-sig")
            return log_path
        except OSError:
            continue
    return None


def unique_output_path(source: Path) -> Path:
    stamp = datetime.now().strftime("%Y%m%d")
    base = source.with_name(f"{source.stem}rw_{stamp}{source.suffix}")
    if not base.exists():
        return base
    index = 2
    while True:
        candidate = source.with_name(
            f"{source.stem}rw_{stamp}_{index:02d}{source.suffix}"
        )
        if not candidate.exists():
            return candidate
        index += 1


def _load_word_automation():
    """Load pywin32, installing it quietly once when Python can use pip."""
    try:
        import pythoncom
        import win32com.client

        return pythoncom, win32com.client
    except ImportError:
        creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        try:
            result = subprocess.run(
                [sys.executable, "-m", "pip", "install", "pywin32"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=creation_flags,
                timeout=180,
                check=False,
            )
            if result.returncode != 0:
                raise RuntimeError
            importlib.invalidate_caches()
            import pythoncom
            import win32com.client

            return pythoncom, win32com.client
        except Exception as exc:
            raise RuntimeError(
                "追蹤修訂或紅色底線功能需要 Windows 版 Microsoft Word 與 pywin32。\n"
                "請確認電腦已安裝 Microsoft Word，並執行：\n"
                "pip install pywin32"
            ) from exc


def _replace_word_hit(
    hit,
    find_value: str,
    replace_value: str,
    red_underline: bool,
):
    """Replace one hit, using Word Find when a merged cell rejects Range.Text."""
    try:
        source_text = str(hit.Text or "")
    except Exception:
        source_text = find_value

    try:
        hit.Text = replace_value
        target = hit
    except Exception as direct_error:
        # A Word table range can contain an internal end-of-cell marker even
        # though the visible match is plain text. Setting Range.Text then fails
        # for vertically/horizontally merged cells. Word's own Find/Replace can
        # replace only the visible match without overwriting the cell marker.
        fallback = hit.Duplicate
        actual_match = re.search(re.escape(find_value), source_text, re.IGNORECASE)
        exact_find_value = actual_match.group(0) if actual_match else find_value
        try:
            finder = fallback.Find
            finder.ClearFormatting()
            finder.Replacement.ClearFormatting()
            if red_underline and replace_value:
                finder.Replacement.Font.Color = 255  # wdColorRed
                finder.Replacement.Font.Underline = 1  # wdUnderlineSingle
            found = finder.Execute(
                FindText=exact_find_value,
                ReplaceWith=replace_value,
                Replace=1,  # wdReplaceOne
                MatchCase=True,
                MatchWholeWord=False,
                MatchWildcards=False,
                Forward=True,
                Wrap=0,  # wdFindStop
                Format=bool(red_underline and replace_value),
            )
            if not found:
                raise RuntimeError("Word 原生替換找不到合併儲存格內的目標文字。")
            target = fallback
        except Exception:
            # Keep the original Word error because it contains the table/cell
            # diagnosis that is most useful in the error log.
            raise direct_error

    if red_underline and replace_value:
        target.Font.Color = 255  # wdColorRed
        target.Font.Underline = 1  # wdUnderlineSingle
    return target


def replace_with_word(
    source: Path,
    output: Path,
    replacements: list[tuple[str, str]],
    track_changes: bool = False,
    red_underline: bool = False,
    progress_callback=None,
) -> int:
    """Read each story once, collect immutable hits, then replace backwards."""
    pythoncom, win32com_client = _load_word_automation()

    pythoncom.CoInitialize()
    word = None
    document = None
    count = 0
    try:
        shutil.copy2(source, output)
        try:
            word = win32com_client.DispatchEx("Word.Application")
        except Exception as exc:
            raise RuntimeError(
                "無法啟動 Microsoft Word。\n"
                "追蹤修訂與紅色底線功能需要已安裝的 Windows 桌面版 Word。"
            ) from exc
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(
            str(output.resolve()),
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False,
        )

        original_track_revisions = bool(document.TrackRevisions)
        original_show_revisions = bool(document.ShowRevisions)
        original_print_revisions = bool(document.PrintRevisions)

        # Searching while Track Changes is active can rediscover deleted text.
        # Freeze the document, collect every original hit, then edit only those
        # stored ranges from back to front. No inserted/deleted revision is ever
        # searched again.
        document.TrackRevisions = False
        document.ShowRevisions = False
        story_hit_groups: list[list[tuple[int, int, int, object, str]]] = []
        total_candidates = 0
        announced_mappings: set[int] = set()
        raw_mapping_counts = [0] * len(replacements)
        search_patterns = [
            re.compile(re.escape(find_value), re.IGNORECASE)
            for find_value, _replace_value in replacements
        ]

        for story_type in range(1, 18):
            try:
                text_range = document.StoryRanges(story_type)
            except Exception:
                continue
            while text_range is not None:
                try:
                    next_story_range = text_range.NextStoryRange
                except Exception:
                    next_story_range = None

                # Word COM calls are comparatively slow. Reading the complete
                # story once and searching in Python avoids one Find.Execute
                # round-trip for every single match.
                story_text = str(text_range.Text or "")
                story_start = int(text_range.Start)
                utf16_offsets = [0]
                for character in story_text:
                    utf16_offsets.append(
                        utf16_offsets[-1] + (2 if ord(character) > 0xFFFF else 1)
                    )
                candidates: list[tuple[int, int, int, object, str]] = []
                for mapping_order, (find_value, replace_value) in enumerate(replacements):
                    if mapping_order not in announced_mappings:
                        announced_mappings.add(mapping_order)
                        _emit_progress(
                            progress_callback,
                            "searching",
                            {"keyword": find_value, "found": 0},
                        )
                    for match in search_patterns[mapping_order].finditer(story_text):
                        match_start, match_end = match.span()
                        hit_start = story_start + utf16_offsets[match_start]
                        hit_end = story_start + utf16_offsets[match_end]
                        hit = text_range.Duplicate
                        hit.SetRange(Start=hit_start, End=hit_end)
                        candidates.append(
                            (hit_start, hit_end, mapping_order, hit, replace_value)
                        )
                        total_candidates += 1
                        raw_mapping_counts[mapping_order] += 1
                        if total_candidates == 1 or total_candidates % 250 == 0:
                            _emit_progress(
                                progress_callback,
                                "searching",
                                {
                                    "keyword": find_value,
                                    "found": raw_mapping_counts[mapping_order],
                                },
                            )
                        if total_candidates > MAX_REPLACEMENTS_PER_DOCUMENT:
                            raise RuntimeError(
                                "單一文件找到超過 100,000 處關鍵字，已停止處理以避免異常。"
                            )

                # Resolve overlapping keywords by the order entered by the user.
                accepted: list[tuple[int, int, int, object, str]] = []
                occupied_until = -1
                for candidate in sorted(
                    candidates, key=lambda item: (item[0], item[2], -(item[1] - item[0]))
                ):
                    if candidate[0] < occupied_until:
                        continue
                    accepted.append(candidate)
                    occupied_until = candidate[1]
                if accepted:
                    story_hit_groups.append(accepted)
                text_range = next_story_range

        keyword_counts = [0] * len(replacements)
        for hits in story_hit_groups:
            for _start, _end, mapping_order, _hit, _replace_value in hits:
                keyword_counts[mapping_order] += 1
        total_hits = sum(keyword_counts)
        _emit_progress(
            progress_callback,
            "search_results",
            {
                "results": [
                    (find_value, keyword_counts[index])
                    for index, (find_value, _replace_value) in enumerate(replacements)
                ],
                "total": total_hits,
            },
        )

        document.TrackRevisions = bool(track_changes)
        if track_changes:
            document.ShowRevisions = True
            document.PrintRevisions = True

        _emit_progress(
            progress_callback,
            "replacing",
            {"current": 0, "total": total_hits, "keyword": ""},
        )
        progress_step = max(1, total_hits // 100) if total_hits else 1
        for hits in story_hit_groups:
            for _start, _end, _order, hit, replace_value in sorted(
                hits, key=lambda item: item[0], reverse=True
            ):
                _replace_word_hit(
                    hit,
                    replacements[_order][0],
                    replace_value,
                    red_underline,
                )
                count += 1
                if count == 1 or count % progress_step == 0 or count == total_hits:
                    _emit_progress(
                        progress_callback,
                        "replacing",
                        {
                            "current": count,
                            "total": total_hits,
                            "keyword": replacements[_order][0],
                        },
                    )

        # If tracking was not requested, preserve all original review settings.
        if not track_changes:
            document.TrackRevisions = original_track_revisions
            document.ShowRevisions = original_show_revisions
            document.PrintRevisions = original_print_revisions
        document.Save()
        document.Close(SaveChanges=False)
        document = None
        return count
    except Exception:
        if output.exists():
            try:
                output.unlink()
            except OSError:
                pass
        raise
    finally:
        if document is not None:
            try:
                document.Close(SaveChanges=False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


_PARAGRAPH_RE = re.compile(
    rb"<(?P<prefix>w|a):p(?:\s[^>]*)?>.*?</(?P=prefix):p>", re.DOTALL
)
_TEXT_NODE_RE = re.compile(
    rb"(<(?P<prefix>w|a):t(?:\s[^>]*)?>)(.*?)(</(?P=prefix):t>)", re.DOTALL
)


def _replace_paragraph_xml(
    paragraph: bytes,
    replacements: list[tuple[str, str]],
    keyword_counts: list[int] | None = None,
    on_change=None,
) -> tuple[bytes, int]:
    matches = list(_TEXT_NODE_RE.finditer(paragraph))
    if not matches:
        return paragraph, 0

    original_node_texts = [html.unescape(m.group(3).decode("utf-8")) for m in matches]
    combined = "".join(original_node_texts)
    boundaries = [0]
    for text_value in original_node_texts:
        boundaries.append(boundaries[-1] + len(text_value))

    candidates: list[tuple[int, int, int]] = []
    for mapping_order, (find_value, _replace_value) in enumerate(replacements):
        for match in re.finditer(re.escape(find_value), combined, re.IGNORECASE):
            start, end = match.span()
            candidates.append((start, end, mapping_order))

    accepted: list[tuple[int, int, int]] = []
    occupied_until = -1
    for candidate in sorted(
        candidates, key=lambda item: (item[0], item[2], -(item[1] - item[0]))
    ):
        if candidate[0] < occupied_until:
            continue
        accepted.append(candidate)
        occupied_until = candidate[1]

    if not accepted:
        return paragraph, 0

    node_texts = list(original_node_texts)
    changed = 0
    for position, end_position, mapping_order in reversed(accepted):
        replace_value = replacements[mapping_order][1]
        start_node = max(0, min(
            len(node_texts) - 1,
            bisect.bisect_right(boundaries, position) - 1,
        ))
        end_node = max(0, min(
            len(node_texts) - 1,
            bisect.bisect_left(boundaries, end_position) - 1,
        ))
        start_offset = position - boundaries[start_node]
        end_offset = end_position - boundaries[end_node]
        if start_node == end_node:
            value = node_texts[start_node]
            node_texts[start_node] = (
                value[:start_offset] + replace_value + value[end_offset:]
            )
        else:
            start_value = node_texts[start_node]
            end_value = node_texts[end_node]
            node_texts[start_node] = start_value[:start_offset] + replace_value
            for index in range(start_node + 1, end_node):
                node_texts[index] = ""
            node_texts[end_node] = end_value[end_offset:]
        changed += 1
        if keyword_counts is not None:
            keyword_counts[mapping_order] += 1
        if on_change is not None:
            on_change(mapping_order)

    chunks: list[bytes] = []
    cursor = 0
    for match, value in zip(matches, node_texts):
        chunks.append(paragraph[cursor : match.start()])
        encoded = html.escape(value, quote=False).encode("utf-8")
        chunks.append(match.group(1) + encoded + match.group(4))
        cursor = match.end()
    chunks.append(paragraph[cursor:])
    return b"".join(chunks), changed


def _replace_xml_part(
    raw_xml: bytes,
    replacements: list[tuple[str, str]],
    keyword_counts: list[int] | None = None,
    on_change=None,
) -> tuple[bytes, int]:
    changed = 0

    def transform(match: re.Match[bytes]) -> bytes:
        nonlocal changed
        value, local_changed = _replace_paragraph_xml(
            match.group(0),
            replacements,
            keyword_counts=keyword_counts,
            on_change=on_change,
        )
        changed += local_changed
        return value

    return _PARAGRAPH_RE.sub(transform, raw_xml), changed


def replace_open_xml(
    source: Path,
    output: Path,
    replacements: list[tuple[str, str]],
    progress_callback=None,
) -> int:
    """Dependency-free fallback for DOCX/DOCM packages."""
    if source.suffix.lower() not in {".docx", ".docm"}:
        raise RuntimeError(
            "舊式 .doc 檔需要安裝 Microsoft Word 與 pywin32，或先另存成 .docx。"
        )
    changed = 0
    try:
        with zipfile.ZipFile(source, "r") as source_zip, zipfile.ZipFile(
            output, "w"
        ) as output_zip:
            infos = source_zip.infolist()
            keyword_counts = [0] * len(replacements)
            for find_value, _replace_value in replacements:
                _emit_progress(
                    progress_callback,
                    "searching",
                    {"keyword": find_value, "found": 0},
                )

            analysis_total = 0

            def report_found(mapping_order: int) -> None:
                nonlocal analysis_total
                analysis_total += 1
                if analysis_total == 1 or analysis_total % 250 == 0:
                    _emit_progress(
                        progress_callback,
                        "searching",
                        {
                            "keyword": replacements[mapping_order][0],
                            "found": keyword_counts[mapping_order],
                        },
                    )

            # Analysis pass uses the exact replacement engine on in-memory data,
            # so the displayed counts match the later write pass.
            for info in infos:
                if info.filename.startswith("word/") and info.filename.endswith(".xml"):
                    data = source_zip.read(info.filename)
                    _replace_xml_part(
                        data,
                        replacements,
                        keyword_counts=keyword_counts,
                        on_change=report_found,
                    )

            total_hits = sum(keyword_counts)
            if total_hits > MAX_REPLACEMENTS_PER_DOCUMENT:
                raise RuntimeError(
                    "單一文件找到超過 100,000 處關鍵字，已停止處理以避免異常。"
                )
            _emit_progress(
                progress_callback,
                "search_results",
                {
                    "results": [
                        (find_value, keyword_counts[index])
                        for index, (find_value, _replace_value) in enumerate(replacements)
                    ],
                    "total": total_hits,
                },
            )
            _emit_progress(
                progress_callback,
                "replacing",
                {"current": 0, "total": total_hits, "keyword": ""},
            )
            progress_step = max(1, total_hits // 100) if total_hits else 1

            def report_change(mapping_order: int) -> None:
                current = report_change.current + 1
                report_change.current = current
                if current == 1 or current % progress_step == 0 or current == total_hits:
                    _emit_progress(
                        progress_callback,
                        "replacing",
                        {
                            "current": current,
                            "total": total_hits,
                            "keyword": replacements[mapping_order][0],
                        },
                    )

            report_change.current = 0
            for info in infos:
                data = source_zip.read(info.filename)
                if info.filename.startswith("word/") and info.filename.endswith(".xml"):
                    data, local_changed = _replace_xml_part(
                        data,
                        replacements,
                        on_change=report_change,
                    )
                    changed += local_changed
                output_zip.writestr(info, data)
        return changed
    except Exception:
        if output.exists():
            try:
                output.unlink()
            except OSError:
                pass
        raise


def _format_inserted_revisions(document) -> None:
    """Apply the requested red underline to every inserted revision."""
    for story_type in range(1, 18):
        try:
            text_range = document.StoryRanges(story_type)
        except Exception:
            continue
        while text_range is not None:
            try:
                next_story_range = text_range.NextStoryRange
            except Exception:
                next_story_range = None
            try:
                revisions = text_range.Revisions
                for index in range(1, int(revisions.Count) + 1):
                    revision = revisions(index)
                    if int(revision.Type) == 1:  # wdRevisionInsert
                        revision.Range.Font.Color = 255  # wdColorRed
                        revision.Range.Font.Underline = 1  # wdUnderlineSingle
            except Exception:
                # A formatting-only failure must not discard a successfully
                # replaced document. Other stories/revisions can still finish.
                pass
            text_range = next_story_range


def replace_with_word_comparison(
    source: Path,
    output: Path,
    replacements: list[tuple[str, str]],
    track_changes: bool,
    red_underline: bool,
    progress_callback=None,
) -> int:
    """Modify Open XML first, then let Word compare documents for revisions."""
    pythoncom, win32com_client = _load_word_automation()
    pythoncom.CoInitialize()
    word = None
    original_document = None
    revised_document = None
    comparison_document = None
    try:
        with tempfile.TemporaryDirectory(prefix="Erod_compare_") as temp_dir:
            revised_path = Path(temp_dir) / f"revised{source.suffix.lower()}"
            count = replace_open_xml(
                source,
                revised_path,
                replacements,
                progress_callback=progress_callback,
            )
            if count == 0:
                shutil.copy2(revised_path, output)
                return 0

            message = (
                "正在由 Word 比較原檔與修改檔，建立追蹤修訂…"
                if track_changes
                else "正在由 Word 標示修改文字…"
            )
            _emit_progress(
                progress_callback,
                "comparing",
                {"message": message},
            )

            try:
                word = win32com_client.DispatchEx("Word.Application")
            except Exception as exc:
                raise RuntimeError(
                    "無法啟動 Microsoft Word，因此無法建立追蹤修訂。"
                ) from exc
            word.Visible = False
            word.DisplayAlerts = 0
            original_document = word.Documents.Open(
                str(source.resolve()),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
            )
            revised_document = word.Documents.Open(
                str(revised_path.resolve()),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
            )

            comparison_document = word.CompareDocuments(
                OriginalDocument=original_document,
                RevisedDocument=revised_document,
                Destination=2,  # wdCompareDestinationNew
                Granularity=0,  # wdGranularityCharLevel
                CompareFormatting=False,
                CompareCaseChanges=True,
                CompareWhitespace=True,
                CompareTables=True,
                CompareHeaders=True,
                CompareFootnotes=True,
                CompareTextboxes=True,
                CompareFields=True,
                CompareComments=False,
                CompareMoves=False,
                RevisedAuthor=APP_NAME,
                IgnoreAllComparisonWarnings=True,
            )

            # Formatting is applied with tracking temporarily disabled so it
            # stays on the inserted text without creating extra format changes.
            comparison_document.TrackRevisions = False
            if red_underline:
                _format_inserted_revisions(comparison_document)
            if track_changes:
                comparison_document.TrackRevisions = True
                comparison_document.ShowRevisions = True
                comparison_document.PrintRevisions = True
            else:
                comparison_document.AcceptAllRevisions()

            file_format = 13 if source.suffix.lower() == ".docm" else 16
            comparison_document.SaveAs2(
                FileName=str(output.resolve()),
                FileFormat=file_format,
                AddToRecentFiles=False,
            )
            comparison_document.Close(SaveChanges=False)
            comparison_document = None
            revised_document.Close(SaveChanges=False)
            revised_document = None
            original_document.Close(SaveChanges=False)
            original_document = None
            _emit_progress(
                progress_callback,
                "comparison_done",
                {"message": "追蹤修訂已建立" if track_changes else "修改標示已完成"},
            )
            return count
    except Exception:
        if output.exists():
            try:
                output.unlink()
            except OSError:
                pass
        raise
    finally:
        for document in (comparison_document, revised_document, original_document):
            if document is not None:
                try:
                    document.Close(SaveChanges=False)
                except Exception:
                    pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def process_document(
    source: Path,
    replacements: list[tuple[str, str]],
    track_changes: bool = False,
    red_underline: bool = False,
    progress_callback=None,
) -> tuple[Path, int]:
    output = unique_output_path(source)
    # DOCX/DOCM text is always modified directly in Open XML, so merged table
    # cells are never edited through Range.Text. Word is used only to compare
    # the original and revised copies when review markup is requested.
    if source.suffix.lower() in {".docx", ".docm"}:
        if track_changes or red_underline:
            count = replace_with_word_comparison(
                source,
                output,
                replacements,
                track_changes=track_changes,
                red_underline=red_underline,
                progress_callback=progress_callback,
            )
        else:
            count = replace_open_xml(
                source,
                output,
                replacements,
                progress_callback=progress_callback,
            )
    else:
        count = replace_with_word(
            source,
            output,
            replacements,
            track_changes=track_changes,
            red_underline=red_underline,
            progress_callback=progress_callback,
        )
    if count == 0:
        try:
            output.unlink()
        except OSError:
            pass
        raise RuntimeError(
            "文件中沒有找到指定的關鍵字，因此未建立輸出檔。\n"
            "請確認文字是否完全相同，並留意全形／半形符號。"
        )
    return output, count


class ErodApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("900x650")
        self.root.minsize(760, 620)
        self.root.configure(bg="#eef2e7")
        self.files: list[Path] = []
        self.output_files: list[Path] = []
        self.last_error_log: Path | None = None
        self.file_status: dict[str, str] = {}
        self.results: queue.Queue[tuple[str, object]] = queue.Queue()
        self.busy = False

        self.find_var = tk.StringVar()
        self.replace_var = tk.StringVar()
        self.track_changes_var = tk.BooleanVar(value=True)
        self.red_underline_var = tk.BooleanVar(value=False)
        self.current_file_var = tk.StringVar(value="目前檔案：尚未開始")
        self.search_results_var = tk.StringVar(value="搜尋結果：等待執行")
        self.replace_progress_var = tk.StringVar(value="已替換：0 / 0")
        self.status_var = tk.StringVar(value="請輸入關鍵字，再拖入 Word 檔案")
        self._configure_style()
        self._build_ui()
        self.root.after_idle(self._fit_initial_window)
        self.root.after(100, self._poll_results)

    def _configure_style(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Toolbar.TFrame", background="#dce6d3")
        style.configure(
            "Option.TCheckbutton",
            background="#dce6d3",
            font=("Microsoft JhengHei UI", 10),
        )
        style.map("Option.TCheckbutton", background=[("active", "#dce6d3")])
        style.configure("Card.TFrame", background="#ffffff")
        style.configure(
            "Work.Horizontal.TProgressbar",
            troughcolor="#d9dfd5",
            background="#47785c",
            bordercolor="#d9dfd5",
            lightcolor="#47785c",
            darkcolor="#47785c",
        )
        style.configure(
            "Batch.Treeview",
            font=("Microsoft JhengHei UI", 10),
            rowheight=28,
            background="#ffffff",
            fieldbackground="#ffffff",
        )
        style.configure(
            "Batch.Treeview.Heading",
            font=("Microsoft JhengHei UI", 10, "bold"),
        )
        style.configure(
            "Primary.TButton",
            font=("Microsoft JhengHei UI", 10, "bold"),
            padding=(16, 8),
            background="#47785c",
            foreground="#ffffff",
        )
        style.map("Primary.TButton", background=[("active", "#365f48")])
        style.configure(
            "Secondary.TButton",
            font=("Microsoft JhengHei UI", 10),
            padding=(12, 8),
        )

    def _build_ui(self) -> None:
        toolbar = ttk.Frame(self.root, style="Toolbar.TFrame", padding=(18, 14))
        toolbar.pack(fill="x")
        toolbar.columnconfigure(1, weight=1)
        toolbar.columnconfigure(3, weight=1)

        label_font = ("Microsoft JhengHei UI", 10, "bold")
        ttk.Label(toolbar, text="尋找關鍵字", font=label_font, background="#dce6d3").grid(
            row=0, column=0, padx=(0, 8), sticky="w"
        )
        find_entry = ttk.Entry(toolbar, textvariable=self.find_var, font=("Microsoft JhengHei UI", 11))
        find_entry.grid(row=0, column=1, padx=(0, 16), sticky="ew", ipady=5)

        ttk.Label(toolbar, text="替換文字", font=label_font, background="#dce6d3").grid(
            row=0, column=2, padx=(0, 8), sticky="w"
        )
        replace_entry = ttk.Entry(
            toolbar, textvariable=self.replace_var, font=("Microsoft JhengHei UI", 11)
        )
        replace_entry.grid(row=0, column=3, padx=(0, 10), sticky="ew", ipady=5)

        ttk.Label(
            toolbar,
            text="多組內容請用「、」分隔；一對一依序替換",
            font=("Microsoft JhengHei UI", 9),
            foreground="#52604f",
            background="#dce6d3",
        ).grid(row=1, column=0, columnspan=4, pady=(8, 0), sticky="w")

        option_frame = ttk.Frame(toolbar, style="Toolbar.TFrame")
        option_frame.grid(row=2, column=0, columnspan=4, pady=(10, 0), sticky="w")
        ttk.Checkbutton(
            option_frame,
            text="建立 Word 追蹤修訂紀錄",
            variable=self.track_changes_var,
            style="Option.TCheckbutton",
        ).pack(side="left")
        ttk.Checkbutton(
            option_frame,
            text="修改文字加紅色底線",
            variable=self.red_underline_var,
            style="Option.TCheckbutton",
        ).pack(side="left", padx=(22, 0))

        progress_panel = tk.Frame(
            self.root,
            bg="#f7f4e8",
            highlightbackground="#b5c2ac",
            highlightthickness=1,
        )
        progress_panel.pack(fill="x", padx=18, pady=(12, 0))
        tk.Label(
            progress_panel,
            textvariable=self.current_file_var,
            font=("Microsoft JhengHei UI", 10, "bold"),
            fg="#293b31",
            bg="#f7f4e8",
            anchor="w",
        ).pack(fill="x", padx=14, pady=(9, 2))
        tk.Label(
            progress_panel,
            textvariable=self.search_results_var,
            font=("Microsoft JhengHei UI", 9),
            fg="#52604f",
            bg="#f7f4e8",
            anchor="w",
            justify="left",
            wraplength=850,
        ).pack(fill="x", padx=14, pady=(0, 5))
        progress_row = tk.Frame(progress_panel, bg="#f7f4e8")
        progress_row.pack(fill="x", padx=14, pady=(0, 9))
        self.replace_progressbar = ttk.Progressbar(
            progress_row,
            orient="horizontal",
            mode="determinate",
            maximum=1,
            value=0,
            style="Work.Horizontal.TProgressbar",
        )
        self.replace_progressbar.pack(side="left", fill="x", expand=True)
        tk.Label(
            progress_row,
            textvariable=self.replace_progress_var,
            font=("Microsoft JhengHei UI", 9, "bold"),
            fg="#405047",
            bg="#f7f4e8",
            width=18,
            anchor="e",
        ).pack(side="right", padx=(12, 0))

        body = ttk.Frame(self.root, style="Card.TFrame", padding=16)
        body.pack(fill="x", padx=18, pady=(18, 10))

        self.drop_zone = tk.Frame(
            body,
            bg="#f7f4e8",
            highlightbackground="#799179",
            highlightthickness=2,
            bd=0,
            cursor="hand2",
        )
        self.drop_zone.pack(fill="x")
        self.drop_zone.bind("<Button-1>", lambda _event: self.choose_files())

        icon = tk.Label(
            self.drop_zone,
            text="W  ↺",
            font=("Segoe UI", 22, "bold"),
            fg="#47785c",
            bg="#f7f4e8",
        )
        icon.pack(side="left", padx=(22, 14), pady=16)
        icon.bind("<Button-1>", lambda _event: self.choose_files())

        drop_text = tk.Frame(self.drop_zone, bg="#f7f4e8")
        drop_text.pack(side="left", fill="both", expand=True, pady=12)
        title = tk.Label(
            drop_text,
            text="拖入多個 Word 檔案加入批次清單",
            font=("Microsoft JhengHei UI", 13, "bold"),
            fg="#293b31",
            bg="#f7f4e8",
            anchor="w",
        )
        title.pack(fill="x")
        title.bind("<Button-1>", lambda _event: self.choose_files())

        subtitle_text = "拖入後只加入清單，不會自動執行；支援 DOC、DOCX、DOCM"
        if not DND_AVAILABLE:
            subtitle_text += "（目前請使用選擇檔案）"
        subtitle = tk.Label(
            drop_text,
            text=subtitle_text,
            font=("Microsoft JhengHei UI", 9),
            fg="#687269",
            bg="#f7f4e8",
            anchor="w",
        )
        subtitle.pack(fill="x", pady=(4, 0))
        subtitle.bind("<Button-1>", lambda _event: self.choose_files())

        self.file_count_label = tk.Label(
            self.drop_zone,
            text="0 個檔案",
            font=("Microsoft JhengHei UI", 11, "bold"),
            fg="#48604e",
            bg="#f7f4e8",
        )
        self.file_count_label.pack(side="right", padx=22)
        self.file_count_label.bind("<Button-1>", lambda _event: self.choose_files())

        if DND_AVAILABLE:
            for target in (
                self.drop_zone,
                icon,
                drop_text,
                title,
                subtitle,
                self.file_count_label,
            ):
                target.drop_target_register(DND_FILES)
                target.dnd_bind("<<Drop>>", self._on_drop)

        list_header = ttk.Frame(body, style="Card.TFrame")
        list_header.pack(fill="x", pady=(12, 7))
        ttk.Label(
            list_header,
            text="批次處理清單",
            font=("Microsoft JhengHei UI", 11, "bold"),
            background="#ffffff",
        ).pack(side="left")
        self.error_log_button = ttk.Button(
            list_header,
            text="開啟錯誤紀錄",
            command=self.open_error_log,
            style="Secondary.TButton",
            state="disabled",
        )
        self.error_log_button.pack(side="right")

        list_frame = ttk.Frame(body, style="Card.TFrame")
        list_frame.pack(fill="both", expand=True)
        self.file_tree = ttk.Treeview(
            list_frame,
            columns=("no", "name", "folder", "status"),
            show="headings",
            selectmode="extended",
            height=5,
            style="Batch.Treeview",
        )
        self.file_tree.heading("no", text="項次")
        self.file_tree.heading("name", text="檔名")
        self.file_tree.heading("folder", text="所在資料夾")
        self.file_tree.heading("status", text="處理狀態")
        self.file_tree.column("no", width=58, minwidth=50, anchor="center", stretch=False)
        self.file_tree.column("name", width=250, minwidth=150, anchor="w")
        self.file_tree.column("folder", width=260, minwidth=150, anchor="w")
        self.file_tree.column("status", width=180, minwidth=130, anchor="center", stretch=False)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=scrollbar.set)
        self.file_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        if DND_AVAILABLE:
            self.file_tree.drop_target_register(DND_FILES)
            self.file_tree.dnd_bind("<<Drop>>", self._on_drop)

        actions = ttk.Frame(self.root, padding=(18, 0, 18, 10))
        actions.pack(fill="x")
        self.run_button = ttk.Button(
            actions,
            text="批次執行替換",
            command=self.start_processing,
            style="Primary.TButton",
        )
        self.run_button.pack(side="right")
        self.add_files_button = ttk.Button(
            actions,
            text="加入檔案",
            command=self.choose_files,
            style="Secondary.TButton",
        )
        self.add_files_button.pack(side="left")
        self.add_folder_button = ttk.Button(
            actions,
            text="加入資料夾",
            command=self.choose_folder,
            style="Secondary.TButton",
        )
        self.add_folder_button.pack(side="left", padx=(8, 0))
        self.remove_button = ttk.Button(
            actions,
            text="移除選取",
            command=self.remove_selected_files,
            style="Secondary.TButton",
        )
        self.remove_button.pack(side="left", padx=(8, 0))
        self.clear_button = ttk.Button(
            actions,
            text="清空清單",
            command=self.clear_files,
            style="Secondary.TButton",
        )
        self.clear_button.pack(side="left", padx=(8, 0))
        ttk.Button(
            actions,
            text="開啟輸出資料夾",
            command=self.open_output_folder,
            style="Secondary.TButton",
        ).pack(side="right", padx=(0, 8))

        footer = tk.Frame(self.root, bg="#e2e6df", height=30)
        footer.pack(fill="x", side="bottom")
        tk.Label(
            footer,
            textvariable=self.status_var,
            font=("Microsoft JhengHei UI", 9),
            fg="#405047",
            bg="#e2e6df",
            anchor="w",
        ).pack(side="left", fill="x", expand=True, padx=12, pady=5)
        tk.Label(
            footer,
            text=VERSION,
            font=("Segoe UI", 9, "bold"),
            fg="#405047",
            bg="#e2e6df",
        ).pack(side="right", padx=12, pady=5)

    def _fit_initial_window(self) -> None:
        """Fit the complete interface inside the Windows desktop work area."""
        self.root.update_idletasks()
        work_left = 0
        work_top = 0
        work_right = self.root.winfo_screenwidth()
        work_bottom = self.root.winfo_screenheight()

        if os.name == "nt":
            try:
                import ctypes
                from ctypes import wintypes

                work_area = wintypes.RECT()
                success = ctypes.windll.user32.SystemParametersInfoW(
                    0x0030,  # SPI_GETWORKAREA: excludes the Windows taskbar
                    0,
                    ctypes.byref(work_area),
                    0,
                )
                if success:
                    work_left = int(work_area.left)
                    work_top = int(work_area.top)
                    work_right = int(work_area.right)
                    work_bottom = int(work_area.bottom)
            except Exception:
                pass

        margin = 12
        available_width = max(760, work_right - work_left - margin * 2)
        available_height = max(560, work_bottom - work_top - margin * 2)

        requested_height = self.root.winfo_reqheight()
        if requested_height > available_height:
            # On high-DPI or smaller screens the file list is the flexible
            # section. Reduce only its visible rows until the footer and all
            # action buttons fit without requiring the user to resize first.
            row_height = 28
            current_rows = int(self.file_tree.cget("height"))
            rows_to_remove = (requested_height - available_height + row_height - 1) // row_height
            self.file_tree.configure(height=max(2, current_rows - rows_to_remove))
            self.root.update_idletasks()

        requested_width = max(900, self.root.winfo_reqwidth())
        requested_height = max(620, self.root.winfo_reqheight())
        width = min(requested_width, available_width)
        height = min(requested_height, available_height)
        x = work_left + max(margin, (work_right - work_left - width) // 2)
        y = work_top + margin
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.minsize(min(760, available_width), min(560, available_height))

    def choose_files(self) -> None:
        if self.busy:
            return
        selected = filedialog.askopenfilenames(
            title="選擇 Word 檔案",
            filetypes=[
                ("Word 檔案", "*.doc *.docx *.docm"),
                ("所有檔案", "*.*"),
            ],
        )
        if selected:
            self._add_files([Path(item) for item in selected])

    def choose_folder(self) -> None:
        if self.busy:
            return
        selected = filedialog.askdirectory(title="選擇含有 Word 檔案的資料夾")
        if not selected:
            return
        folder = Path(selected)
        paths = sorted(
            (
                path
                for path in folder.iterdir()
                if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS
            ),
            key=lambda path: path.name.lower(),
        )
        if not paths:
            messagebox.showinfo("沒有 Word 檔案", "所選資料夾內沒有支援的 Word 檔案。")
            return
        self._add_files(paths)

    def _on_drop(self, event: tk.Event) -> None:
        if self.busy:
            return
        paths = [Path(item) for item in self.root.tk.splitlist(event.data)]
        expanded: list[Path] = []
        for path in paths:
            if path.is_dir():
                expanded.extend(
                    sorted(
                        (
                            item
                            for item in path.iterdir()
                            if item.is_file()
                            and item.suffix.lower() in SUPPORTED_EXTENSIONS
                        ),
                        key=lambda item: item.name.lower(),
                    )
                )
            else:
                expanded.append(path)
        # Drag and drop is queue-only. Processing is intentionally reachable
        # only through the "批次執行替換" button's start_processing command.
        self._add_files(expanded)

    @staticmethod
    def _file_key(path: Path) -> str:
        return os.path.normcase(str(path.resolve()))

    def _add_files(self, paths: list[Path]) -> None:
        valid: list[Path] = []
        invalid: list[str] = []
        existing = {self._file_key(path) for path in self.files}
        duplicates = 0
        for path in paths:
            if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS:
                key = self._file_key(path)
                if key in existing:
                    duplicates += 1
                    continue
                existing.add(key)
                valid.append(path)
            else:
                invalid.append(path.name)
        self.files.extend(valid)
        if valid:
            for path in valid:
                self.file_status[self._file_key(path)] = "等待處理"
            self._refresh_file_tree()
            message = (
                f"已加入 {len(valid)} 個檔案，目前共 {len(self.files)} 個；"
                "請按「批次執行替換」開始處理"
            )
            if duplicates:
                message += f"；略過 {duplicates} 個重複檔案"
            self.status_var.set(message)
        if invalid:
            messagebox.showwarning(
                "不支援的檔案",
                "以下檔案不是支援的 Word 格式：\n" + "\n".join(invalid),
            )

    def _refresh_file_tree(self) -> None:
        selected_keys = set(self.file_tree.selection())
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        for index, path in enumerate(self.files, start=1):
            key = self._file_key(path)
            self.file_tree.insert(
                "",
                "end",
                iid=key,
                values=(
                    index,
                    path.name,
                    str(path.parent),
                    self.file_status.get(key, "等待處理"),
                ),
            )
            if key in selected_keys:
                self.file_tree.selection_add(key)
        self.file_count_label.config(text=f"{len(self.files)} 個檔案")

    def remove_selected_files(self) -> None:
        if self.busy:
            return
        selected = set(self.file_tree.selection())
        if not selected:
            messagebox.showinfo("尚未選取", "請先在批次清單中選取要移除的檔案。")
            return
        self.files = [path for path in self.files if self._file_key(path) not in selected]
        for key in selected:
            self.file_status.pop(key, None)
        self._refresh_file_tree()
        self.status_var.set(f"已移除 {len(selected)} 個檔案")

    def clear_files(self) -> None:
        if self.busy:
            return
        self.files.clear()
        self.output_files.clear()
        self.last_error_log = None
        self.error_log_button.config(state="disabled")
        self.file_status.clear()
        self._refresh_file_tree()
        self.status_var.set("批次清單已清空")

    def start_processing(self) -> None:
        if self.busy:
            return
        if not self.files:
            messagebox.showinfo("尚未載入", "請先拖入或選擇 Word 檔案。")
            return
        try:
            replacements = make_replacements(self.find_var.get(), self.replace_var.get())
        except ValueError as exc:
            messagebox.showwarning("關鍵字設定", str(exc))
            return

        missing = [str(path) for path in self.files if not path.exists()]
        if missing:
            messagebox.showerror("檔案不存在", "找不到：\n" + "\n".join(missing))
            return

        files = list(self.files)
        track_changes = self.track_changes_var.get()
        red_underline = self.red_underline_var.get()
        self.output_files = []
        self.last_error_log = None
        self.error_log_button.config(state="disabled")
        self.busy = True
        self._set_controls_enabled(False)
        self.current_file_var.set("目前檔案：準備開始")
        self.search_results_var.set("搜尋結果：等待搜尋")
        self.replace_progress_var.set("已替換：0 / 0")
        self.replace_progressbar.config(maximum=1, value=0)
        for path in files:
            self.file_status[self._file_key(path)] = "等待處理"
        self._refresh_file_tree()
        self.status_var.set(f"正在處理 {len(files)} 個檔案，請稍候…")
        threading.Thread(
            target=self._worker,
            args=(files, replacements, track_changes, red_underline),
            daemon=True,
        ).start()

    def _worker(
        self,
        files: list[Path],
        replacements: list[tuple[str, str]],
        track_changes: bool,
        red_underline: bool,
    ) -> None:
        failed: list[tuple[Path, str, str, str]] = []
        batch_completed: list[tuple[Path, Path, int]] = []
        for index, path in enumerate(files, start=1):
            self.results.put(("progress", (path, "處理中", index, len(files))))
            try:
                def report_document_progress(stage, data, current_path=path) -> None:
                    self.results.put(
                        (
                            "document_progress",
                            (current_path, stage, data, index, len(files)),
                        )
                    )

                output, count = process_document(
                    path,
                    replacements,
                    track_changes=track_changes,
                    red_underline=red_underline,
                    progress_callback=report_document_progress,
                )
                batch_completed.append((path, output, count))
                self.results.put(
                    ("progress", (path, f"完成（{count} 處）", index, len(files)))
                )
            except Exception as exc:
                technical = traceback.format_exc()
                summary, suggestion = explain_error(str(exc))
                failed.append((path, summary, suggestion, technical))
                self.results.put(
                    ("progress", (path, f"失敗：{summary}", index, len(files)))
                )
        error_log = write_error_log(
            files,
            replacements,
            track_changes,
            red_underline,
            batch_completed,
            failed,
        )
        self.results.put(("finished", (batch_completed, failed, error_log)))

    def _poll_results(self) -> None:
        try:
            while True:
                event_name, payload = self.results.get_nowait()
                if event_name == "progress":
                    path, status, index, total = payload
                    self._set_tree_status(path, status)
                    self.status_var.set(f"批次處理中：{index}/{total}　{path.name}")
                elif event_name == "document_progress":
                    self._handle_document_progress(*payload)
                elif event_name == "finished":
                    self._finish_processing(payload)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_results)

    def _set_tree_status(self, path: Path, status: str) -> None:
        key = self._file_key(path)
        self.file_status[key] = status
        if self.file_tree.exists(key):
            values = list(self.file_tree.item(key, "values"))
            values[3] = status
            self.file_tree.item(key, values=values)
            self.file_tree.see(key)

    def _handle_document_progress(
        self,
        path: Path,
        stage: str,
        data: object,
        batch_index: int,
        batch_total: int,
    ) -> None:
        details = data if isinstance(data, dict) else {}
        self.current_file_var.set(
            f"目前檔案：{path.name}　（批次 {batch_index}/{batch_total}）"
        )
        if stage == "searching":
            keyword = str(details.get("keyword", ""))
            found = int(details.get("found", 0) or 0)
            suffix = f"，目前找到 {found} 處" if found else ""
            self.search_results_var.set(f"正在搜尋：「{keyword}」{suffix}")
            self._set_tree_status(path, f"搜尋中：{keyword}")
            self.status_var.set(
                f"正在搜尋關鍵字　批次 {batch_index}/{batch_total}　{path.name}"
            )
        elif stage == "search_results":
            results = details.get("results", [])
            total = int(details.get("total", 0) or 0)
            result_text = "；".join(
                f"「{keyword}」{count} 處" for keyword, count in results
            )
            self.search_results_var.set(
                f"搜尋結果：{result_text}　｜　合計 {total} 處"
            )
            self.replace_progressbar.config(maximum=max(total, 1), value=0)
            self.replace_progress_var.set(f"已替換：0 / {total}")
            self._set_tree_status(path, f"找到 {total} 處")
        elif stage == "replacing":
            current = int(details.get("current", 0) or 0)
            total = int(details.get("total", 0) or 0)
            keyword = str(details.get("keyword", ""))
            self.replace_progressbar.config(
                maximum=max(total, 1), value=min(current, max(total, 1))
            )
            self.replace_progress_var.set(f"已替換：{current} / {total}")
            self._set_tree_status(path, f"替換 {current}/{total}")
            keyword_text = f"「{keyword}」" if keyword else "文字"
            self.status_var.set(
                f"正在替換{keyword_text}　{current}/{total}　"
                f"批次 {batch_index}/{batch_total}"
            )
        elif stage == "comparing":
            message = str(
                details.get("message", "正在由 Word 建立追蹤修訂…")
            )
            self.replace_progress_var.set("正在建立修訂…")
            self._set_tree_status(path, "建立追蹤修訂中")
            self.status_var.set(
                f"{message}　批次 {batch_index}/{batch_total}　{path.name}"
            )
        elif stage == "comparison_done":
            message = str(details.get("message", "追蹤修訂已建立"))
            self.replace_progress_var.set(message)
            self._set_tree_status(path, message)
            self.status_var.set(
                f"{message}　批次 {batch_index}/{batch_total}　{path.name}"
            )

    def _finish_processing(
        self,
        payload: tuple[
            list[tuple[Path, Path, int]],
            list[tuple[Path, str, str, str]],
            Path | None,
        ],
    ) -> None:
        completed, failed, error_log = payload
        self.busy = False
        self._set_controls_enabled(True)
        self.last_error_log = error_log
        self.error_log_button.config(state="normal" if error_log else "disabled")
        if completed:
            total = sum(count for _source, _output, count in completed)
            self.output_files = [output for _source, output, _count in completed]
            self.current_file_var.set(
                f"批次處理完成：成功 {len(completed)} 個，失敗 {len(failed)} 個"
            )
            self.status_var.set(
                f"完成 {len(completed)} 個檔案，共替換 {total} 處文字"
            )
            self.replace_progress_var.set(f"處理完成：{total} 處")
        else:
            self.current_file_var.set("批次處理結束：沒有成功輸出的檔案")
            self.status_var.set("處理失敗，未建立輸出檔")

        if failed:
            details = "\n\n".join(
                f"{path.name}\n原因：{summary}\n建議：{suggestion}"
                for path, summary, suggestion, _technical in failed
            )
            if error_log:
                details += f"\n\n完整錯誤紀錄：\n{error_log}"
            messagebox.showerror("部分檔案處理失敗", details)
        elif completed:
            messagebox.showinfo(
                "替換完成",
                f"已建立 {len(completed)} 個新檔案。\n"
                f"共替換 {sum(count for _source, _output, count in completed)} 處文字。",
            )

    def _set_controls_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        for button in (
            self.run_button,
            self.add_files_button,
            self.add_folder_button,
            self.remove_button,
            self.clear_button,
        ):
            button.config(state=state)

    def open_error_log(self) -> None:
        if not self.last_error_log or not self.last_error_log.exists():
            messagebox.showinfo("尚無錯誤紀錄", "目前沒有可開啟的錯誤紀錄檔。")
            return
        try:
            if os.name == "nt":
                os.startfile(str(self.last_error_log))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(self.last_error_log)])
            else:
                subprocess.Popen(["xdg-open", str(self.last_error_log)])
        except Exception as exc:
            messagebox.showerror("無法開啟錯誤紀錄", str(exc))

    def open_output_folder(self) -> None:
        if not self.files:
            messagebox.showinfo("尚無檔案", "請先載入或輸出 Word 檔案。")
            return
        folder = (
            self.output_files[0].parent
            if self.output_files
            else self.files[0].parent
        )
        try:
            if os.name == "nt":
                os.startfile(str(folder))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(folder)])
            else:
                subprocess.Popen(["xdg-open", str(folder)])
        except Exception as exc:
            messagebox.showerror("無法開啟資料夾", str(exc))


def main() -> None:
    root = TkinterDnD.Tk() if DND_AVAILABLE else tk.Tk()
    ErodApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
