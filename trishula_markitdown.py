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
from typing import List, Dict, Optional, Set

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
            # Strip trailing/leading spaces, escape pipes, and replace inner newlines
            clean_cell = cell.strip().replace('|', '\\|')
            clean_cell = re.sub(r'\s+', ' ', clean_cell)
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

def get_sheet_paths(z: zipfile.ZipFile) -> List[tuple[str, str]]:
    """Parse workbook.xml and workbook.xml.rels to get list of (sheet_name, worksheet_path)."""
    try:
        workbook_xml = z.read("xl/workbook.xml")
    except KeyError:
        return [("Sheet1", "xl/worksheets/sheet1.xml")]

    root = ET.fromstring(workbook_xml)
    ns = {
        "ns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    }

    # 1. Parse sheets
    sheets = []
    sheet_nodes = root.findall(".//ns:sheet", ns)
    if not sheet_nodes:
        for child in root.iter():
            if child.tag.endswith("}sheet") or child.tag == "sheet":
                name = child.attrib.get("name", "")
                r_id = child.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
                if name and r_id:
                    sheets.append((name, r_id))
    else:
        for node in sheet_nodes:
            name = node.attrib.get("name", "")
            r_id = node.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
            if name and r_id:
                sheets.append((name, r_id))

    if not sheets:
        return [("Sheet1", "xl/worksheets/sheet1.xml")]

    # 2. Parse workbook.xml.rels to resolve sheet relationship paths
    try:
        rels_xml = z.read("xl/_rels/workbook.xml.rels")
    except KeyError:
        return [(name, f"xl/worksheets/sheet{i+1}.xml") for i, (name, _) in enumerate(sheets)]

    rels_root = ET.fromstring(rels_xml)
    rels_map = {}
    for child in rels_root.iter():
        if child.tag.endswith("}Relationship") or child.tag == "Relationship":
            rel_id = child.attrib.get("Id", "")
            target = child.attrib.get("Target", "")
            if rel_id and target:
                rels_map[rel_id] = target

    # Map sheets to workbook-relative paths
    sheet_paths = []
    for name, r_id in sheets:
        path = rels_map.get(r_id, "")
        if path:
            if not path.startswith("xl/"):
                path = "xl/" + path
            sheet_paths.append((name, path))

    if not sheet_paths:
        return [("Sheet1", "xl/worksheets/sheet1.xml")]
    return sheet_paths

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
    """Parse .xlsx Excel sheet into Markdown tables, supporting multiple worksheets."""
    if not zipfile.is_zipfile(file_path):
        raise ValueError("Invalid Excel file: Not a ZIP archive")

    with zipfile.ZipFile(file_path) as z:
        shared_strings = get_shared_strings(z)
        sheet_paths = get_sheet_paths(z)
        
        sheet_markdowns = []
        for sheet_name, sheet_path in sheet_paths:
            try:
                sheet_content = z.read(sheet_path)
            except KeyError:
                continue

            root = ET.fromstring(sheet_content)
            sheet_data = None
            for child in root.iter():
                if child.tag.endswith("}sheetData") or child.tag == "sheetData":
                    sheet_data = child
                    break

            if sheet_data is None:
                continue

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
                            elif cell_type == "b":
                                val = "TRUE" if val_str == "1" else "FALSE"
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

            table_md = format_markdown_table(rows)
            if table_md:
                if len(sheet_paths) > 1:
                    sheet_markdowns.append(f"## Sheet: {sheet_name}\n" + table_md)
                else:
                    sheet_markdowns.append(table_md)

        return "\n\n".join(sheet_markdowns).strip()

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
        is_strike = False
        is_underline = False
        rPr = None
        for child in r_node:
            if child.tag.endswith("}rPr") or child.tag == "rPr":
                rPr = child
                break
        if rPr is not None:
            for prop in rPr:
                tag = prop.tag
                if tag.endswith("}b") or tag == "b":
                    is_bold = True
                elif tag.endswith("}i") or tag == "i":
                    is_italic = True
                elif tag.endswith("}strike") or tag == "strike":
                    is_strike = True
                elif tag.endswith("}u") or tag == "u":
                    is_underline = True

        parts = []
        for child in r_node:
            tag = child.tag
            if tag.endswith("}t") or tag == "t":
                if child.text:
                    parts.append(child.text)
            elif tag.endswith("}br") or tag == "br":
                parts.append("  \n")
            elif tag.endswith("}drawing") or tag == "drawing":
                for desc in child.iter():
                    if desc.tag.endswith("}blip") or desc.tag == "blip":
                        embed_id = desc.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed") or desc.attrib.get("embed", "")
                        if embed_id:
                            img_path = rels.get(embed_id, "")
                            if img_path:
                                img_name = os.path.basename(img_path)
                                parts.append(f"![Image]({img_name})")

        text = "".join(parts)
        if text:
            if is_bold:
                text = f"**{text}**"
            if is_italic:
                text = f"*{text}*"
            if is_strike:
                text = f"~~{text}~~"
            if is_underline:
                text = f"<u>{text}</u>"
        return text

    def render_paragraph(p_node: ET.Element) -> tuple[bool, str]:
        """Render paragraph blocks resolving runs, nested hyperlinks, and list items."""
        is_list_item = False
        ilvl = 0
        pPr = None
        for child in p_node:
            if child.tag.endswith("}pPr") or child.tag == "pPr":
                pPr = child
                break
        heading_level = 0
        if pPr is not None:
            pStyle = None
            for child in pPr:
                if child.tag.endswith("}pStyle") or child.tag == "pStyle":
                    pStyle = child
                    break
            if pStyle is not None:
                val = pStyle.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") or pStyle.attrib.get("val", "")
                match = re.match(r'(?i)Heading\s*([1-6])', val)
                if match:
                    heading_level = int(match.group(1))

            numPr = None
            for child in pPr:
                if child.tag.endswith("}numPr") or child.tag == "numPr":
                    numPr = child
                    break
            if numPr is not None:
                is_list_item = True
                for child in numPr:
                    if child.tag.endswith("}ilvl") or child.tag == "ilvl":
                        try:
                            ilvl = int(child.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "0"))
                        except ValueError:
                            ilvl = 0

        runs = []
        for child in p_node:
            tag = child.tag
            if tag.endswith("}r") or tag == "r":
                runs.append(render_run(child))
            elif tag.endswith("}hyperlink") or tag == "hyperlink":
                r_id = child.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
                anchor = child.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}anchor") or child.attrib.get("anchor", "")
                url = rels.get(r_id, "")
                link_text = "".join(render_run(run_child) for run_child in child if run_child.tag.endswith("}r") or run_child.tag == "r")
                if url:
                    runs.append(f"[{link_text}]({url})")
                elif anchor:
                    runs.append(f"[{link_text}](#{anchor})")
                else:
                    runs.append(link_text)
        
        text = "".join(runs).strip()
        if text:
            if is_list_item:
                indent = "  " * ilvl
                text = f"{indent}- {text}"
            elif heading_level > 0:
                text = "#" * heading_level + " " + text
        return is_list_item, text

    def render_table(tbl_node: ET.Element) -> str:
        rows = []
        for tr in tbl_node:
            if not (tr.tag.endswith("}tr") or tr.tag == "tr"):
                continue
            row_cells = []
            for tc in tr:
                if tc.tag.endswith("}tc") or tc.tag == "tc":
                    grid_span = 1
                    tcPr = None
                    for child in tc:
                        if child.tag.endswith("}tcPr") or child.tag == "tcPr":
                            tcPr = child
                            break
                    if tcPr is not None:
                        for child in tcPr:
                            if child.tag.endswith("}gridSpan") or child.tag == "gridSpan":
                                try:
                                    grid_span = int(child.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") or child.attrib.get("val", "1"))
                                except ValueError:
                                    grid_span = 1

                    cell_paragraphs = []
                    for p in tc:
                        if p.tag.endswith("}p") or p.tag == "p":
                            _, p_text = render_paragraph(p)
                            cell_paragraphs.append(p_text)
                    cell_val = " ".join(cell_paragraphs).strip()
                    row_cells.append(cell_val)
                    for _ in range(grid_span - 1):
                        row_cells.append("")
            if row_cells:
                rows.append(row_cells)
        return format_markdown_table(rows)

    list_items = []
    for child in body:
        tag = child.tag
        if tag.endswith("}p") or tag == "p":
            is_list, text = render_paragraph(child)
            if text:
                if is_list:
                    list_items.append(text)
                else:
                    if list_items:
                        markdown_blocks.append("\n".join(list_items))
                        list_items = []
                    markdown_blocks.append(text)
        elif tag.endswith("}tbl") or tag == "tbl":
            if list_items:
                markdown_blocks.append("\n".join(list_items))
                list_items = []
            table_md = render_table(child)
            if table_md:
                markdown_blocks.append(table_md)

    if list_items:
        markdown_blocks.append("\n".join(list_items))

    return "\n\n".join(markdown_blocks).strip()

def detect_csv_delimiter(file_path: str) -> str:
    """Heuristic to detect CSV delimiter (, or ; or \t) from first line."""
    try:
        with open(file_path, "r", encoding="utf-8", errors="replace") as f:
            first_line = f.readline()
        if not first_line:
            return ","
        counts = {
            ",": first_line.count(","),
            ";": first_line.count(";"),
            "\t": first_line.count("\t")
        }
        best_delim = max(counts, key=counts.get)
        if counts[best_delim] > 0:
            return best_delim
    except:
        pass
    return ","

def parse_csv(file_path: str) -> str:
    """Parse .csv file into a formatted Markdown table, dynamically detecting delimiter."""
    delimiter = detect_csv_delimiter(file_path)
    rows = []
    with open(file_path, "r", newline="", encoding="utf-8", errors="replace") as f:
        reader = csv.reader(f, delimiter=delimiter)
        for row in reader:
            rows.append(row)
    # Trim trailing empty rows
    while rows and all(not cell.strip() for cell in rows[-1]):
        rows.pop()
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
