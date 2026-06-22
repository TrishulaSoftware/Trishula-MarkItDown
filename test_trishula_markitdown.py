import unittest
import os
from unittest.mock import patch
import zipfile
import io
import tempfile
import csv

from trishula_markitdown import (
    col_letter_to_index, parse_cell_ref, format_markdown_table,
    parse_csv, parse_xlsx, parse_docx, convert_to_markdown
)

class TestUtilities(unittest.TestCase):
    def test_col_letter_to_index(self):
        self.assertEqual(col_letter_to_index("A"), 0)
        self.assertEqual(col_letter_to_index("B"), 1)
        self.assertEqual(col_letter_to_index("Z"), 25)
        self.assertEqual(col_letter_to_index("AA"), 26)
        self.assertEqual(col_letter_to_index("AZ"), 51)
        self.assertEqual(col_letter_to_index("BA"), 52)

    def test_parse_cell_ref(self):
        self.assertEqual(parse_cell_ref("A1"), (0, 0))
        self.assertEqual(parse_cell_ref("B12"), (11, 1))
        self.assertEqual(parse_cell_ref("AA100"), (99, 26))

    def test_format_markdown_table(self):
        rows = [
            ["Runner", "Odds"],
            ["Blue Storm", "+150"],
            ["Old Yeller", "+300"]
        ]
        md = format_markdown_table(rows)
        self.assertIn("| Runner     | Odds |", md)
        self.assertIn("| ---------- | ---- |", md)
        self.assertIn("| Blue Storm | +150 |", md)
        self.assertIn("| Old Yeller | +300 |", md)

class TestCSVParser(unittest.TestCase):
    def setUp(self):
        self.temp_file = tempfile.NamedTemporaryFile(suffix=".csv", delete=False, mode="w", encoding="utf-8")
        writer = csv.writer(self.temp_file)
        writer.writerow(["Name", "Points", "Rebounds"])
        writer.writerow(["LeBron James", "25", "7"])
        writer.writerow(["Anthony Davis", "22", "12"])
        self.temp_file.close()

    def tearDown(self):
        os.unlink(self.temp_file.name)

    def test_csv_conversion(self):
        md = parse_csv(self.temp_file.name)
        self.assertIn("| Name          | Points | Rebounds |", md)
        self.assertIn("| LeBron James  | 25     | 7        |", md)
        self.assertIn("| Anthony Davis | 22     | 12       |", md)

class TestXLSXParser(unittest.TestCase):
    def setUp(self):
        # Create a mock zip-based Excel sheet in-memory
        self.temp_xlsx = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        self.temp_xlsx.close()
        
        shared_strings_xml = (
            '<?xml version="1.0" encoding="utf-8"?>'
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">'
            '<si><t>Player</t></si>'
            '<si><t>Spread</t></si>'
            '<si><t>Team A</t></si>'
            '</sst>'
        )

        sheet1_xml = (
            '<?xml version="1.0" encoding="utf-8"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetData>'
            # Row 1 has cell refs
            '<row r="1">'
            '<c r="A1" t="s"><v>0</v></c>'
            '<c r="B1" t="s"><v>1</v></c>'
            '</row>'
            # Row 2 tests sequential fallbacks (no row ref) and inlineStr cells
            '<row>'
            '<c t="s"><v>2</v></c>'
            '<c t="inlineStr"><is><t>+4.5</t></is></c>'
            '</row>'
            '</sheetData>'
            '</worksheet>'
        )

        with zipfile.ZipFile(self.temp_xlsx.name, "w") as z:
            z.writestr("xl/sharedStrings.xml", shared_strings_xml)
            z.writestr("xl/worksheets/sheet1.xml", sheet1_xml)

    def tearDown(self):
        os.unlink(self.temp_xlsx.name)

    def test_xlsx_conversion(self):
        md = parse_xlsx(self.temp_xlsx.name)
        # Check standard cells
        self.assertIn("| Player | Spread |", md)
        # Check sequential row parsing + inlineStr
        self.assertIn("| Team A | +4.5   |", md)

class TestDOCXParser(unittest.TestCase):
    def setUp(self):
        # Create a mock zip-based Word document in-memory
        self.temp_docx = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        self.temp_docx.close()

        document_rels_xml = (
            '<?xml version="1.0" encoding="utf-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://trishula.local"/>'
            '</Relationships>'
        )

        document_xml = (
            '<?xml version="1.0" encoding="utf-8"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<w:body>'
            '<w:p>'
            '<w:r>'
            '<w:rPr><w:b/></w:rPr>'
            '<w:t>Sovereign System Briefing</w:t>'
            '</w:r>'
            '</w:p>'
            # Paragraph with nested hyperlinks and breaks
            '<w:p>'
            '<w:r>'
            '<w:t>This is line 1.</w:t>'
            '<w:br/>'
            '<w:t>Line 2 leads to </w:t>'
            '</w:r>'
            '<w:hyperlink r:id="rId1">'
            '<w:r>'
            '<w:t>Trishula Home</w:t>'
            '</w:r>'
            '</w:hyperlink>'
            '</w:p>'
            '<w:tbl>'
            '<w:tr>'
            '<w:tc>'
            '<w:p><w:r><w:t>Col 1</w:t></w:r></w:p>'
            '</w:tc>'
            '<w:tc>'
            '<w:p><w:r><w:t>Col 2</w:t></w:r></w:p>'
            '</w:tc>'
            '</w:tr>'
            '</w:tbl>'
            '</w:body>'
            '</w:document>'
        )

        with zipfile.ZipFile(self.temp_docx.name, "w") as z:
            z.writestr("word/document.xml", document_xml)
            z.writestr("word/_rels/document.xml.rels", document_rels_xml)

    def tearDown(self):
        os.unlink(self.temp_docx.name)

    def test_docx_conversion(self):
        md = parse_docx(self.temp_docx.name)
        # Check bold text rendering
        self.assertIn("**Sovereign System Briefing**", md)
        # Check paragraph breaks (double spaces)
        self.assertIn("This is line 1.  \nLine 2 leads to", md)
        # Check resolved hyperlinks
        self.assertIn("[Trishula Home](http://trishula.local)", md)
        # Check table rendering
        self.assertIn("| Col 1 | Col 2 |", md)

class TestFormatRouter(unittest.TestCase):
    @patch("os.path.exists")
    def test_unsupported_format(self, mock_exists):
        mock_exists.return_value = True
        with self.assertRaises(ValueError):
            convert_to_markdown("test_file.pdf")

if __name__ == "__main__":
    unittest.main()
