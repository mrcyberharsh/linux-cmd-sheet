"""Tests for the DOCX command-sheet search utility."""
import tempfile
import unittest
from pathlib import Path

from docx import Document

from search_sheet import list_sections, search_commands


class SearchSheetTests(unittest.TestCase):
    def test_searches_paragraphs_and_tables(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "sample.docx"
            doc = Document()
            doc.add_heading("Networking", level=1)
            doc.add_paragraph("ip addr show")
            table = doc.add_table(rows=1, cols=2)
            table.cell(0, 0).text = "ss -tulpen"
            table.cell(0, 1).text = "Show listening sockets"
            doc.save(path)
            self.assertEqual(list_sections(path), ["Networking"])
            results = search_commands(path, "listening")
            self.assertEqual(len(results), 1)
            self.assertIn("ss -tulpen", results[0]["text"])

    def test_rejects_empty_query(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "sample.docx"
            Document().save(path)
            with self.assertRaises(ValueError):
                search_commands(path, " ")


if __name__ == "__main__":
    unittest.main()
