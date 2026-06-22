import unittest
import os
from unittest.mock import patch
import zipfile
import io
import tempfile
import csv

from trishula_markitdown import (
    col_letter_to_index, parse_cell_ref, format_markdown_table,
    parse_csv, parse_xlsx, parse_docx, convert_to_markdown, detect_csv_delimiter
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

    def test_format_markdown_table_escapes_pipes(self):
        rows = [
            ["Runner | Team", "Odds"],
            ["Blue | Storm", "+150"]
        ]
        md = format_markdown_table(rows)
        self.assertIn("Runner \\| Team", md)
        self.assertIn("Blue \\| Storm", md)

class TestCSVParser(unittest.TestCase):
    def setUp(self):
        self.temp_csv = tempfile.NamedTemporaryFile(suffix=".csv", delete=False, mode="w", encoding="utf-8")
        self.temp_csv.write("Name;Points;Rebounds\nLeBron James;25;7\nAnthony Davis;22;12\n;;;\n\n")
        self.temp_csv.close()

    def tearDown(self):
        os.unlink(self.temp_csv.name)

    def test_csv_delimiter_detection(self):
        # Verify semicolon is correctly detected
        delim = detect_csv_delimiter(self.temp_csv.name)
        self.assertEqual(delim, ";")

    def test_csv_conversion(self):
        # Convert semi-colon delimited file
        md = parse_csv(self.temp_csv.name)
        self.assertIn("| Name          | Points | Rebounds |", md)
        self.assertIn("| LeBron James  | 25     | 7        |", md)
        self.assertIn("| Anthony Davis | 22     | 12       |", md)
        # Ensure trailing empty rows are trimmed
        self.assertNotIn("|               |        |          |", md)

class TestXLSXParser(unittest.TestCase):
    def setUp(self):
        # Create a mock zip-based Excel sheet in-memory
        self.temp_xlsx = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        self.temp_xlsx.close()
        
        shared_strings_xml = (
            '<?xml version="1.0" encoding="utf-8"?>'
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">'
            '<si><t>Player</t></si>'
            '<si><t>Spread</t></si>'
            '<si><t>Team A</t></si>'
            '<si><t>Market</t></si>'
            '</sst>'
        )

        workbook_xml = (
            '<?xml version="1.0" encoding="utf-8"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets>'
            '<sheet name="OddsSheet" sheetId="1" r:id="rId1"/>'
            '<sheet name="InfoSheet" sheetId="2" r:id="rId2"/>'
            '</sheets>'
            '</workbook>'
        )

        workbook_rels_xml = (
            '<?xml version="1.0" encoding="utf-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
            '</Relationships>'
        )

        sheet1_xml = (
            '<?xml version="1.0" encoding="utf-8"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetData>'
            '<row r="1">'
            '<c r="A1" t="s"><v>0</v></c>'
            '<c r="B1" t="s"><v>1</v></c>'
            '</row>'
            '<row>'
            '<c t="s"><v>2</v></c>'
            '<c t="inlineStr"><is><t>+4.5</t></is></c>'
            '</row>'
            '<row>'
            '<c t="b"><v>1</v></c>'
            '<c t="b"><v>0</v></c>'
            '</row>'
            '</sheetData>'
            '</worksheet>'
        )

        sheet2_xml = (
            '<?xml version="1.0" encoding="utf-8"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetData>'
            '<row r="1">'
            '<c r="A1" t="s"><v>3</v></c>'
            '</row>'
            '<row>'
            '<c t="inlineStr"><is><t>Pinnacle</t></is></c>'
            '</row>'
            '</sheetData>'
            '</worksheet>'
        )

        with zipfile.ZipFile(self.temp_xlsx.name, "w") as z:
            z.writestr("xl/sharedStrings.xml", shared_strings_xml)
            z.writestr("xl/workbook.xml", workbook_xml)
            z.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
            z.writestr("xl/worksheets/sheet1.xml", sheet1_xml)
            z.writestr("xl/worksheets/sheet2.xml", sheet2_xml)

    def tearDown(self):
        os.unlink(self.temp_xlsx.name)

    def test_xlsx_conversion(self):
        md = parse_xlsx(self.temp_xlsx.name)
        # Check Sheet 1 output and header
        self.assertIn("## Sheet: OddsSheet", md)
        self.assertIn("| Player | Spread |", md)
        self.assertIn("| Team A | +4.5   |", md)
        self.assertIn("| TRUE   | FALSE  |", md)
        # Check Sheet 2 output and header
        self.assertIn("## Sheet: InfoSheet", md)
        self.assertIn("| Market   |", md)
        self.assertIn("| Pinnacle |", md)

class TestDOCXParser(unittest.TestCase):
    def setUp(self):
        # Create a mock zip-based Word document in-memory
        self.temp_docx = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        self.temp_docx.close()

        document_rels_xml = (
            '<?xml version="1.0" encoding="utf-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://trishula.local"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>'
            '</Relationships>'
        )

        document_xml = (
            '<?xml version="1.0" encoding="utf-8"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<w:body>'
            '<w:p>'
            '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
            '<w:r><w:t>Sovereign System Briefing</w:t></w:r>'
            '</w:p>'
            # List item 1 (w:ilvl = 0)
            '<w:p>'
            '<w:pPr>'
            '<w:numPr>'
            '<w:ilvl w:val="0"/>'
            '<w:numId w:val="1"/>'
            '</w:numPr>'
            '</w:pPr>'
            '<w:r><w:t>Bullet Level 0</w:t></w:r>'
            '</w:p>'
            # List item 2 (w:ilvl = 1)
            '<w:p>'
            '<w:pPr>'
            '<w:numPr>'
            '<w:ilvl w:val="1"/>'
            '<w:numId w:val="1"/>'
            '</w:numPr>'
            '</w:pPr>'
            '<w:r><w:t>Bullet Level 1</w:t></w:r>'
            '</w:p>'
            # Strikethrough & Underline
            '<w:p>'
            '<w:r>'
            '<w:rPr><w:strike/></w:rPr>'
            '<w:t>Deleted Text</w:t>'
            '</w:r>'
            '<w:r>'
            '<w:rPr><w:u w:val="single"/></w:rPr>'
            '<w:t>Underlined Text</w:t>'
            '</w:r>'
            '</w:p>'
            # Drawing/Image
            '<w:p>'
            '<w:r>'
            '<w:drawing>'
            '<a:blip r:embed="rId2"/>'
            '</w:drawing>'
            '</w:r>'
            '</w:p>'
            # Hyperlink & Anchor
            '<w:p>'
            '<w:hyperlink r:id="rId1">'
            '<w:r><w:t>Trishula Home</w:t></w:r>'
            '</w:hyperlink>'
            '<w:hyperlink w:anchor="section2">'
            '<w:r><w:t>Go to Sec 2</w:t></w:r>'
            '</w:hyperlink>'
            '</w:p>'
            # Table with spans & pipes
            '<w:tbl>'
            '<w:tr>'
            '<w:tc>'
            '<w:p><w:r><w:t>Header A | B</w:t></w:r></w:p>'
            '</w:tc>'
            '<w:tc>'
            '<w:p><w:r><w:t>Header C</w:t></w:r></w:p>'
            '</w:tc>'
            '</w:tr>'
            '<w:tr>'
            '<w:tc>'
            '<w:tcPr><w:gridSpan w:val="2"/></w:tcPr>'
            '<w:p><w:r><w:t>Span Cell</w:t></w:r></w:p>'
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
        # Check heading style
        self.assertIn("# Sovereign System Briefing", md)
        # Check bullet item rendering and indentation
        self.assertIn("- Bullet Level 0\n  - Bullet Level 1", md)
        # Check strike & underline styles
        self.assertIn("~~Deleted Text~~", md)
        self.assertIn("<u>Underlined Text</u>", md)
        # Check image parsing
        self.assertIn("![Image](image1.png)", md)
        # Check resolved hyperlinks
        self.assertIn("[Trishula Home](http://trishula.local)", md)
        # Check bookmark anchor link
        self.assertIn("[Go to Sec 2](#section2)", md)
        # Check table pipe escaping and cell merging (span)
        self.assertIn("Header A \\| B", md)
        self.assertIn("Span Cell", md)

class TestFormatRouter(unittest.TestCase):
    @patch("os.path.exists")
    def test_unsupported_format(self, mock_exists):
        mock_exists.return_value = True
        with self.assertRaises(ValueError):
            convert_to_markdown("test_file.pdf")

if __name__ == "__main__":
    unittest.main()
