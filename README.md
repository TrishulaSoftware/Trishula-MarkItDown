# ⚜️ Trishula-MarkItDown
> **Sovereign Document-to-Markdown Parser Core**

A lightweight, zero-dependency, pure-Python document-to-markdown converter built entirely within the Python standard library. Designed to offer local, air-gapped parsing of Excel (`.xlsx`), Word (`.docx`), and CSV documents without external package installations or network egress risks.

Optimized for sports analytics, odds sheets, and ledger feeds where tabular layouts, formulas, text styles, and document structures must be converted cleanly into model-readable markdown format.

---

## █ Strategic Alignment & Features
* **Zero Dependencies**: Pure Python implementation using only standard library modules (`xml.etree.ElementTree`, `zipfile`, `csv`, `re`, `io`). Zero threat of supply-chain attacks.
* **Excel (`.xlsx`) Engine**: Resolves sparse row grids, shared string indices (`sharedStrings.xml`), inline text (`inlineStr`), numbers, booleans, and formats tables using standard markdown grid layouts.
* **Word (`.docx`) Engine**: Converts structured paragraphs, text formatting runs (bold, italic, strike), list structures, line breaks, tables (`w:tbl`), and extracts links.
* **CSV Engine**: Safely reads CSV records with proper quote escaping and formats them into clean, aligned markdown tables.
* **CLI Command Shell**: Exposes clean stdin/stdout piping and output file redirection options.

---

## █ Installation & Requirements
* **Runtime Environment**: Python 3.10, 3.11, or 3.12 (standard library).
* **Installation**: Drop `trishula_markitdown.py` directly into your project root. No `pip install` required.

---

## █ Usage Reference

### Convert Excel Odds Sheet to Markdown
```bash
python trishula_markitdown.py path/to/betting_odds.xlsx
```

### Convert Word Briefing to Markdown
```bash
python trishula_markitdown.py path/to/briefing.docx
```

### Convert CSV ledger and output to a file
```bash
python trishula_markitdown.py path/to/ledger.csv -o output.md
```

---

## █ Proof of Work (Verified Console Output)

The suite has been verified using a self-contained mock ZIP/XML generation mechanism validating Office Open XML specifications:

```
> python test_trishula_markitdown.py
.........
----------------------------------------------------------------------
Ran 9 tests in 0.062s

OK
```

---

## █ CI/CD Integration
This repository is configured with a GitHub Actions workflow (`.github/workflows/ci.yml`) validating changes against Python versions `3.10`, `3.11`, and `3.12` on every push to the `main` branch.
