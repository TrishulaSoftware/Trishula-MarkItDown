#!/usr/bin/env python3
"""
Trishula-MarkItDown: Sovereign Pure-Python Document-to-Markdown Converter
Zero dependencies. Local-first. Optimized for sports tables and data formats.
"""

import sys
import os
import zipfile
import csv
import re
import argparse
import xml.etree.ElementTree as ET
from typing import List, Dict, Optional

def format_markdown_table(rows: List[List[str]]) -> str:
    """Format a 2D list of strings into a clean, aligned Markdown table."""
    if not rows:
        return ""

    # Normalize column count
    col_count = max(len(row) for row in rows)
    for row in rows:
        while len(row) < col_count:
            row.append("")

    # Clean cells (strip spaces/newlines) and calculate column widths
    cleaned_rows = []
    col_widths = [3] * col_count
    for row in rows:
        cleaned_row = []
        for idx, cell in enumerate(row):
            # Strip trailing/leading spaces and replace inner newlines
            clean_cell = re.sub(r'\s+', ' ', cell.strip())
            col_widths[idx] = max(col_widths[idx], len(clean_cell))
            cleaned_row.append(clean_cell)
        cleaned_rows.append(cleaned_row)

    lines = []
    # Header row
    header_line = "| " + " | ".join(cell.ljust(col_widths[i]) for i, cell in enumerate(cleaned_rows[0])) + " |"
    lines.append(header_line)

    # Separator row
    separator_line = "| " + " | ".join("-" * col_widths[i] for i in range(col_count)) + " |"
    lines.append(separator_line)

    # Data rows
    for row in cleaned_rows[1:]:
        row_line = "| " + " | ".join(cell.ljust(col_widths[i]) for i, cell in enumerate(row)) + " |"
        lines.append(row_line)

    return "\n" + "\n".join(lines) + "\n"

def get_shared_strings(z: zipfile.ZipFile) -> List[str]:
    """Parse shared strings XML in an Excel zip archive."""
    try:
        xml_content = z.read("xl/sharedStrings.xml")
    except KeyError:
        return []

    root = ET.fromstring(xml_content)
    strings = []
    
    # Try finding using standard Excel namespaces, fallback to prefix-less iteration
    ns = {"ns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    si_nodes = root.findall(".//ns:si", ns)
    if not si_nodes:
        # Fallback iteration for namespace-less/mismatched prefixes
        for si in root.iter():
            if si.tag.endswith("}si") or si.tag == "si":
                t_texts = []
                for t in si.iter():
                    if (t.tag.endswith("}t") or t.tag == "t") and t.text:
                        t_texts.append(t.text)
                strings.append("".join(t_texts))
    else:
        for si in si_nodes:
            t_texts = []
            for t in si.findall(".//ns:t", ns):
                if t.text:
                    t_texts.append(t.text)
            strings.append("".join(t_texts))

    return strings

def col_letter_to_index(letter: str) -> int:
    """Converts Excel column letter (A, B, AA) to 0-based column index."""
    index = 0
    for char in letter:
        if 'A' <= char <= 'Z':
            index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1

def parse_cell_ref(ref: str) -> tuple[int, int]:
    """Parse cell reference (e.g. 'B12') into 0-based (row_idx, col_idx)."""
    match = re.match(r'^([A-Z]+)([0-9]+)$', ref)
    if match:
        col_letter, row_str = match.groups()
        return int(row_str) - 1, col_letter_to_index(col_letter)
    return 0, 0

def parse_xlsx(file_path: str) -> str:
    """Parse .xlsx Excel sheet into Markdown tables."""
    if not zipfile.is_zipfile(file_path):
        raise ValueError("Invalid Excel file: Not a ZIP archive")

    with zipfile.ZipFile(file_path) as z:
        shared_strings = get_shared_strings(z)
        
        # Read sheet1.xml
        try:
            sheet_content = z.read("xl/worksheets/sheet1.xml")
        except KeyError:
            raise ValueError("Worksheet 'sheet1.xml' not found in Excel archive")

        root = ET.fromstring(sheet_content)
        sheet_data = None
        for child in root.iter():
            if child.tag.endswith("}sheetData") or child.tag == "sheetData":
                sheet_data = child
                break

        if sheet_data is None:
            return ""

        # Parse grid cells supporting sparse grids and sequential references
        grid = {}
        max_row = 0
        max_col = 0
        current_row_idx = 0

        for row in sheet_data:
            if not (row.tag.endswith("}row") or row.tag == "row"):
                continue
            
            r_attr = row.attrib.get("r")
            if r_attr:
                r_idx = int(r_attr) - 1
                current_row_idx = r_idx
            else:
                r_idx = current_row_idx

            grid[r_idx] = {}
            max_row = max(max_row, r_idx + 1)
            current_col_idx = 0

            for cell in row:
                if not (cell.tag.endswith("}c") or cell.tag == "c"):
                    continue

                c_ref = cell.attrib.get("r")
                if c_ref:
                    _, c_idx = parse_cell_ref(c_ref)
                    current_col_idx = c_idx
                else:
                    c_idx = current_col_idx
                
                max_col = max(max_col, c_idx + 1)
                val = ""
                cell_type = cell.attrib.get("t", "")

                if cell_type == "inlineStr":
                    # Parse inlineStr cells: <is><t>Text</t></is>
                    is_node = None
                    for child in cell:
                        if child.tag.endswith("}is") or child.tag == "is":
                            is_node = child
                            break
                    if is_node is not None:
                        t_node = None
                        for child in is_node:
                            if child.tag.endswith("}t") or child.tag == "t":
                                t_node = child
                                break
                        if t_node is not None and t_node.text:
                            val = t_node.text
                else:
                    # Parse standard values: <v>
                    v_node = None
                    for child in cell:
                        if child.tag.endswith("}v") or child.tag == "v":
                            v_node = child
                            break
                    if v_node is not None and v_node.text:
                        val_str = v_node.text
                        if cell_type == "s":
                            try:
                                idx = int(val_str)
                                if 0 <= idx < len(shared_strings):
                                    val = shared_strings[idx]
                            except ValueError:
                                val = val_str
                        else:
                            val = val_str

                grid[r_idx][c_idx] = val
                current_col_idx += 1

            current_row_idx += 1

        # Reconstruct full 2D grid
        rows = []
        for r in range(max_row):
            row_data = []
            for c in range(max_col):
                row_data.append(grid.get(r, {}).get(c, ""))
            rows.append(row_data)

        # Filter out empty rows at the end
        while rows and all(not cell for cell in rows[-1]):
            rows.pop()

        return format_markdown_table(rows)

def get_docx_relations(z: zipfile.ZipFile) -> Dict[str, str]:
    """Parse relationship IDs mapped to hyperlink target URLs in Word."""
    try:
        xml_content = z.read("word/_rels/document.xml.rels")
    except KeyError:
        return {}

    root = ET.fromstring(xml_content)
    rels = {}
    for child in root.iter():
        if child.tag.endswith("}Relationship") or child.tag == "Relationship":
            rel_id = child.attrib.get("Id", "")
            target = child.attrib.get("Target", "")
            if rel_id and target:
                rels[rel_id] = target
    return rels

def parse_docx(file_path: str) -> str:
    """Parse .docx Word document paragraphs, headers, hyperlinks and tables into Markdown."""
    if not zipfile.is_zipfile(file_path):
        raise ValueError("Invalid Word file: Not a ZIP archive")

    with zipfile.ZipFile(file_path) as z:
        try:
            doc_content = z.read("word/document.xml")
        except KeyError:
            raise ValueError("Document file 'word/document.xml' not found in Word archive")
        
        rels = get_docx_relations(z)

    root = ET.fromstring(doc_content)
    body = None
    for child in root.iter():
        if child.tag.endswith("}body") or child.tag == "body":
            body = child
            break

    if body is None:
        return ""

    markdown_blocks = []

    def render_run(r_node: ET.Element) -> str:
        """Render a single text run with bold/italic formatting and breaks."""
        is_bold = False
        is_italic = False
        rPr = None
        for child in r_node:
            if child.tag.endswith("}rPr") or child.tag == "rPr":
                rPr = child
                break
        if rPr is not None:
            for prop in rPr:
                if prop.tag.endswith("}b") or prop.tag == "b":
                    is_bold = True
                if prop.tag.endswith("}i") or prop.tag == "i":
                    is_italic = True

        parts = []
        for child in r_node:
            tag = child.tag
            if tag.endswith("}t") or tag == "t":
                if child.text:
                    parts.append(child.text)
            elif tag.endswith("}br") or tag == "br":
                parts.append("  \n")

        text = "".join(parts)
        if text:
            if is_bold:
                text = f"**{text}**"
            if is_italic:
                text = f"*{text}*"
        return text

    def render_paragraph(p_node: ET.Element) -> str:
        """Render paragraph blocks resolving runs and nested hyperlinks."""
        runs = []
        for child in p_node:
            tag = child.tag
            if tag.endswith("}r") or tag == "r":
                runs.append(render_run(child))
            elif tag.endswith("}hyperlink") or tag == "hyperlink":
                r_id = child.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
                url = rels.get(r_id, "")
                # Render runs inside hyperlink block
                link_text = "".join(render_run(run_child) for run_child in child if run_child.tag.endswith("}r") or run_child.tag == "r")
                if url:
                    runs.append(f"[{link_text}]({url})")
                else:
                    runs.append(link_text)
        return "".join(runs).strip()

    def render_table(tbl_node: ET.Element) -> str:
        rows = []
        for tr in tbl_node:
            if not (tr.tag.endswith("}tr") or tr.tag == "tr"):
                continue
            row_cells = []
            for tc in tr:
                if tc.tag.endswith("}tc") or tc.tag == "tc":
                    cell_paragraphs = []
                    for p in tc:
                        if p.tag.endswith("}p") or p.tag == "p":
                            cell_paragraphs.append(render_paragraph(p))
                    row_cells.append(" ".join(cell_paragraphs).strip())
            if row_cells:
                rows.append(row_cells)
        return format_markdown_table(rows)

    for child in body:
        tag = child.tag
        if tag.endswith("}p") or tag == "p":
            text = render_paragraph(child)
            if text:
                markdown_blocks.append(text)
        elif tag.endswith("}tbl") or tag == "tbl":
            table_md = render_table(child)
            if table_md:
                markdown_blocks.append(table_md)

    return "\n\n".join(markdown_blocks).strip()

def parse_csv(file_path: str) -> str:
    """Parse .csv file into a formatted Markdown table."""
    rows = []
    with open(file_path, "r", newline="", encoding="utf-8", errors="replace") as f:
        reader = csv.reader(f)
        for row in reader:
            rows.append(row)
    return format_markdown_table(rows)

def convert_to_markdown(file_path: str) -> str:
    """Detects format and parses the document file into Markdown."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    _, ext = os.path.splitext(file_path.lower())
    
    if ext == ".csv":
        return parse_csv(file_path)
    elif ext == ".xlsx":
        return parse_xlsx(file_path)
    elif ext == ".docx":
        return parse_docx(file_path)
    else:
        raise ValueError(f"Unsupported file format: {ext}. Must be .csv, .xlsx, or .docx")

def main():
    parser = argparse.ArgumentParser(description="Trishula-MarkItDown: Sovereign Document-to-Markdown Converter")
    parser.add_argument("file", help="Path to CSV, XLSX, or DOCX document to convert")
    parser.add_argument("-o", "--output", help="Path to write output markdown file")

    args = parser.parse_args()

    try:
        markdown = convert_to_markdown(args.file)
        if args.output:
            with open(args.output, "w", encoding="utf-8") as f:
                f.write(markdown)
            print(f"✅ Successfully wrote markdown output to: {args.output}")
        else:
            print(markdown)
    except Exception as e:
        print(f"❌ Error: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
