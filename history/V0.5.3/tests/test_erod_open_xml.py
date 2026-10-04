from __future__ import annotations

import importlib.util
import sys
import tempfile
import unittest
import zipfile
from importlib.machinery import SourceFileLoader
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = ROOT / "Erod_V0.5.3.py"


def load_erod():
    loader = SourceFileLoader("erod_v053", str(MODULE_PATH))
    spec = importlib.util.spec_from_loader(loader.name, loader)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


class OpenXmlReplacementTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.erod = load_erod()

    def test_replacements_do_not_cascade_into_later_keywords(self) -> None:
        paragraph = (
            b"<w:p>"
            b"<w:r><w:t>A</w:t></w:r>"
            b"<w:r><w:t> B</w:t></w:r>"
            b"</w:p>"
        )
        updated, changed = self.erod._replace_paragraph_xml(
            paragraph,
            [("A", "B"), ("B", "C")],
        )
        self.assertEqual(changed, 2)
        self.assertIn(b"<w:t>B</w:t>", updated)
        self.assertIn(b"<w:t> C</w:t>", updated)
        self.assertNotIn(b"<w:t>C</w:t></w:r><w:r><w:t> C</w:t>", updated)

    def test_split_text_nodes_can_be_replaced_as_one_visible_word(self) -> None:
        paragraph = (
            "<w:p>"
            "<w:r><w:t>承包</w:t></w:r>"
            "<w:r><w:t>商</w:t></w:r>"
            "</w:p>"
        ).encode("utf-8")
        updated, changed = self.erod._replace_paragraph_xml(
            paragraph,
            [("承包商", "廠商")],
        )
        self.assertEqual(changed, 1)
        self.assertIn("廠商".encode("utf-8"), updated)

    def test_earlier_overlapping_keyword_wins(self) -> None:
        paragraph = "<w:p><w:r><w:t>工程司</w:t></w:r></w:p>".encode("utf-8")
        updated, changed = self.erod._replace_paragraph_xml(
            paragraph,
            [("工程", "施工"), ("工程司", "監造單位")],
        )
        self.assertEqual(changed, 1)
        self.assertIn("施工司".encode("utf-8"), updated)
        self.assertNotIn("監造單位".encode("utf-8"), updated)

    def test_replace_open_xml_reports_original_text_counts(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            source = temp_path / "source.docx"
            output = temp_path / "output.docx"
            document_xml = (
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                b"<w:document>"
                b"<w:body>"
                b"<w:p><w:r><w:t>A</w:t></w:r><w:r><w:t> B</w:t></w:r></w:p>"
                b"</w:body>"
                b"</w:document>"
            )
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("word/document.xml", document_xml)
                archive.writestr("docProps/core.xml", b"<core/>")

            progress_events = []

            def progress(stage, data):
                if stage == "search_results":
                    progress_events.append(data)

            changed = self.erod.replace_open_xml(
                source,
                output,
                [("A", "B"), ("B", "C")],
                progress_callback=progress,
            )

            self.assertEqual(changed, 2)
            self.assertTrue(output.exists())
            self.assertEqual(progress_events[-1]["total"], 2)
            self.assertEqual(progress_events[-1]["results"], [("A", 1), ("B", 1)])
            with zipfile.ZipFile(output) as archive:
                updated_xml = archive.read("word/document.xml")
            self.assertIn(b"<w:t>B</w:t>", updated_xml)
            self.assertIn(b"<w:t> C</w:t>", updated_xml)

    def test_process_document_falls_back_without_word(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            source = temp_path / "source.docx"
            document_xml = (
                b'<w:document><w:body><w:p><w:r><w:t>A</w:t></w:r>'
                b"</w:p></w:body></w:document>"
            )
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("word/document.xml", document_xml)

            events = []
            original_status = self.erod.word_automation_status
            self.erod.word_automation_status = lambda: (False, "沒有 Word")
            try:
                output, changed = self.erod.process_document(
                    source,
                    [("A", "B")],
                    track_changes=True,
                    red_underline=True,
                    progress_callback=lambda stage, data: events.append((stage, data)),
                )
            finally:
                self.erod.word_automation_status = original_status

            self.assertEqual(changed, 1)
            self.assertTrue(output.exists())
            self.assertIn("fallback", [stage for stage, _data in events])
            with zipfile.ZipFile(output) as archive:
                self.assertIn(b"<w:t>B</w:t>", archive.read("word/document.xml"))


if __name__ == "__main__":
    unittest.main()
