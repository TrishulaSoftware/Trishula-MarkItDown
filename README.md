# Trishula-MarkItDown

A lightweight, zero-dependency document-to-markdown converter built purely in Python's standard library. Designed to offer local parsing of Excel (`.xlsx`), Word (`.docx`), and CSV documents without external package installations or security egress.

Optimized for sports statistics, odds tables, and projection feeds where grid alignments and nested document hierarchies must be preserved.

---

## █ Features
* **Zero Dependencies**: Uses only standard Python library modules (`xml.etree`, `zipfile`, `csv`, `re`).
* **Excel (.xlsx) Parsing**: Resolves cell coordinates, sparse row fallbacks, shared strings lookup tables, and inline strings (`inlineStr`).
* **Word (.docx) Parsing**: Parses paragraphs, nested runs (bold/italic styles), line breaks (`w:br`), tables (`w:tbl`), and maps relationship IDs to resolve hyperlinks.
* **CSV Parsing**: Converts text-based comma-separated grids into clean, aligned Markdown tables.
* **CLI-Ready**: Accepts a file path and outputs the converted Markdown or writes it directly to a file.

---

## █ Installation & Requirements
* **Requirements**: Python 3.10+ (pure standard library).
* **Installation**: Drop `trishula_markitdown.py` into your path or project root.

---

## █ Usage Reference

### 1. Convert Excel sheet to Markdown Table
```bash
python trishula_markitdown.py path/to/betting_odds.xlsx
```

### 2. Convert Word Document to Markdown
```bash
python trishula_markitdown.py path/to/briefing.docx
```

### 3. Convert CSV and Save to File
```bash
python trishula_markitdown.py path/to/ledger.csv -o output.md
```

---

## █ Running Tests
To run the mock-based unit test suite (validating OpenXML schemas and ZIP decompression):
```bash
python -m unittest test_trishula_markitdown.py
```
