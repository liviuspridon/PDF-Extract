#!/usr/bin/env python3
"""
Extracts the 5 tables from Mengenkalkulation PDFs using word-level
X/Y positioning — robust across different machines and fonts.

Usage: python3 extract_tables.py input.pdf [output_dir]
"""

import sys
import re
import csv
import pdfplumber
from collections import defaultdict


def get_words(pdf_path):
    """Extract all words with their positions across all pages."""
    words = []
    with pdfplumber.open(pdf_path) as pdf:
        y_offset = 0
        for page in pdf.pages:
            page_words = page.extract_words() or []
            for w in page_words:
                words.append({
                    "text": w["text"],
                    "x0": w["x0"],
                    "x1": w["x1"],
                    "top": w["top"] + y_offset
                })
            y_offset += page.height
    return words


def group_by_row(words, tolerance=3):
    """Group words into rows by proximity of their top coordinate."""
    rows = defaultdict(list)
    for w in words:
        # find existing row within tolerance
        matched = None
        for key in rows:
            if abs(key - w["top"]) <= tolerance:
                matched = key
                break
        if matched is None:
            matched = w["top"]
        rows[matched].append(w)
    # sort each row by x position
    return {k: sorted(rows[k], key=lambda w: w["x0"]) for k in sorted(rows)}


def find_section_top(rows, section_name):
    """Find the Y coordinate where a section header appears."""
    for top, words in rows.items():
        text = " ".join(w["text"] for w in words)
        if section_name in text:
            return top
    return None


def words_in_band(rows, y_start, y_end, x_min=0, x_max=9999):
    """Return rows of words within a vertical and horizontal band."""
    result = {}
    for top, words in rows.items():
        if y_start < top < y_end:
            filtered = [w for w in words if x_min <= w["x0"] <= x_max]
            if filtered:
                result[top] = filtered
    return result


def merge_text(words):
    """Merge a list of word dicts into a single string."""
    return " ".join(w["text"] for w in words)


def is_number(s):
    return bool(re.match(r"^-?\d+([.,]\d+)?$", s))


def norm(s):
    return s.replace(",", ".")


def parse_data_rows(rows, num_value_cols, y_start, y_end):
    """
    Parse data rows within a vertical band.
    Each row starts with an integer (Pos), ends with num_value_cols numbers,
    and everything in between is the material name.
    """
    result = []
    for top, words in sorted(rows.items()):
        if top <= y_start or top >= y_end:
            continue
        texts = [w["text"] for w in words]
        if not texts or not re.match(r"^\d+$", texts[0]):
            continue
        pos = texts[0]
        rest = texts[1:]
        if len(rest) < num_value_cols:
            continue
        values = rest[-num_value_cols:]
        if not all(is_number(v) for v in values):
            continue
        material = " ".join(rest[:-num_value_cols])
        result.append([pos, material] + [norm(v) for v in values])
    return result


def extract_tables(pdf_path, out_dir="."):
    words = get_words(pdf_path)
    rows = group_by_row(words)

    tables = {}

    # ── 1. Plattenmaterial ──────────────────────────────────────────
    y1 = find_section_top(rows, "Plattenmaterial")
    y2 = find_section_top(rows, "Material pentru margini")
    if y1 and y2:
        headers = ["Pos", "Material", "Starke", "Stuck", "Lange", "Breite",
                   "Brutto_qm", "Brutto_pct", "Netto_qm", "Netto_pct",
                   "Standard_qm", "Standard_pct", "EK_qm", "EK_pct",
                   "AS_qm", "AS_pct", "VK_qm", "Summe"]
        data = parse_data_rows(rows, 16, y1, y2)
        tables["Plattenmaterial"] = [headers] + data

    # ── 2. Material pentru margini ──────────────────────────────────
    y3 = find_section_top(rows, "Taiere")
    if y2 and y3:
        headers = ["Pos", "Material", "Inaltime", "Bucata", "Lungime",
                   "Net_m", "Net_pct", "Brut_m", "Brut_pct",
                   "Standard_m", "Standard_pct", "EK_m", "EK_pct", "Total"]
        data = parse_data_rows(rows, 13, y2, y3)
        tables["Material_pentru_margini"] = [headers] + data

    # ── 3. Taiere ───────────────────────────────────────────────────
    y4 = find_section_top(rows, "Bandarea marginilor")
    if y3 and y4:
        headers = ["Pos", "Material", "Grosime", "metri", "metri_per_qm",
                   "Lungime_per_Stck", "Zeit_min", "Pret_Std",
                   "Pret_metri", "Pret_qm", "Total"]
        data = parse_data_rows(rows, 9, y3, y4)
        tables["Taiere"] = [headers] + data

    # ── 4. Bandarea marginilor ──────────────────────────────────────
    y5 = find_section_top(rows, "Costuri speciale")
    if y4 and y5:
        headers = ["Pos", "Material", "Grosime", "Inaltime", "Lungime_metri",
                   "per_Stck", "Bucata", "Zeit_min", "Pret_Std",
                   "Pret_metri", "AS_pct", "PE_metri", "Total"]
        data = parse_data_rows(rows, 11, y4, y5)
        tables["Bandarea_marginilor"] = [headers] + data

    # ── 5. Costuri speciale ─────────────────────────────────────────
    y6 = find_section_top(rows, "Zwischen")
    if y5 and y6:
        headers = ["Pos", "Cantitate", "EP", "Tip", "Total"]
        data = []
        for top, row_words in sorted(rows.items()):
            if top <= y5 or top >= y6:
                continue
            texts = [w["text"] for w in row_words]
            if not texts or not re.match(r"^\d+$", texts[0]):
                continue
            if len(texts) < 3:
                continue
            pos = texts[0]
            total = texts[-1]
            cant = texts[1]
            ep = texts[2]
            tip = " ".join(texts[3:-1]) if len(texts) > 4 else texts[3] if len(texts) > 3 else ""
            data.append([pos, norm(cant), norm(ep), tip, norm(total)])
        tables["Costuri_speciale"] = [headers] + data

    # ── Write CSVs ──────────────────────────────────────────────────
    for name, table_rows in tables.items():
        out_path = f"{out_dir}/{name}.csv"
        with open(out_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerows(table_rows)
        print(f"Wrote {out_path} ({len(table_rows)-1} data rows)")

    return tables


if __name__ == "__main__":
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else "input.pdf"
    out_dir = sys.argv[2] if len(sys.argv) > 2 else "."
    extract_tables(pdf_path, out_dir)
