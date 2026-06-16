#!/usr/bin/env python3
"""
Extracts the 5 tables from this specific 'Mengenkalkulation' PDF type
and writes them to CSV files (one CSV per table, plus a combined CSV).

Tables:
  1. Plattenmaterial      -> 18 numeric fields after material name
  2. Material pentru margini -> 12 numeric fields after material name
  3. Taiere               -> 8 numeric fields after material name
  4. Bandarea marginilor  -> 10 numeric fields after material name
  5. Costuri speciale     -> Pos, Cantitate, EP, Tip, Total

Usage: python3 extract_tables.py input.pdf output_dir
"""

import sys
import re
import csv
import pdfplumber


def get_full_text(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            text += page_text + "\n"
    # remove page-break markers
    text = re.sub(r"\nSeite \d+\n", "\n", text)
    return text


def parse_numeric_rows(lines, start_idx, num_value_cols, headers):
    """
    Parse rows starting with an integer Pos number, followed by a
    material name (variable words) and exactly num_value_cols numeric values.
    Stops at the first line that is just a single number (table total)
    or doesn't match the pattern.
    Returns (rows, next_index)
    """
    rows = [headers]
    i = start_idx
    # regex: leading pos number, then capture rest
    row_re = re.compile(r"^(\d+)\s+(.*)$")
    num_re = re.compile(r"^-?\d+([.,]\d+)?$")

    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        m = row_re.match(line)
        if not m:
            break
        pos = m.group(1)
        rest = m.group(2).split()
        if len(rest) < num_value_cols:
            break
        values = rest[-num_value_cols:]
        # validate all values look numeric
        if not all(num_re.match(v) for v in values):
            break
        material = " ".join(rest[:-num_value_cols])
        # normalize decimal commas to dots
        values_norm = [v.replace(",", ".") for v in values]
        rows.append([pos, material] + values_norm)
        i += 1

    # skip the table-total line (single number) if present
    if i < len(lines):
        total_line = lines[i].strip()
        if re.match(r"^-?\d+([.,]\d+)?$", total_line):
            i += 1

    return rows, i


def parse_costuri_speciale(lines, start_idx):
    """Parse the 'Costuri speciale' table: Pos, Cantitate, EP, Tip, Total"""
    rows = [["Pos", "Cantitate", "EP", "Tip", "Total"]]
    i = start_idx
    row_re = re.compile(r"^(\d+)\s+(\S+)\s+(\S+)\s+(.+)\s+(\S+)$")
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        m = row_re.match(line)
        if not m:
            break
        pos, cant, ep, tip, total = m.groups()
        rows.append([pos, cant.replace(",", "."), ep.replace(",", "."), tip, total.replace(",", ".")])
        i += 1
    # skip total line
    if i < len(lines) and re.match(r"^-?\d+([.,]\d+)?$", lines[i].strip()):
        i += 1
    return rows, i


def main(pdf_path, out_dir="."):
    text = get_full_text(pdf_path)
    lines = text.split("\n")

    tables = {}

    # find "Plattenmaterial" section
    for idx, line in enumerate(lines):
        if line.startswith("Plattenmaterial"):
            start = idx + 2  # skip 2 header lines
            headers = ["Pos", "Material", "Starke", "Stuck", "Lange", "Breite",
                       "Brutto_qm", "Brutto_pct", "Netto_qm", "Netto_pct",
                       "Standard_qm", "Standard_pct", "EK_qm", "EK_pct",
                       "AS_qm", "AS_pct", "VK_qm", "Summe"]
            rows, _ = parse_numeric_rows(lines, start, 16, headers)
            tables["Plattenmaterial"] = rows
            break

    # Material pentru margini
    for idx, line in enumerate(lines):
        if line.startswith("Material pentru margini"):
            start = idx + 2
            headers = ["Pos", "Material", "Inaltime", "Bucata", "Lungime",
                       "Net_m", "Net_pct", "Brut_m", "Brut_pct",
                       "Standard_m", "Standard_pct", "EK_m", "EK_pct", "Total"]
            rows, _ = parse_numeric_rows(lines, start, 13, headers)
            tables["Material_pentru_margini"] = rows
            break

    # Taiere
    for idx, line in enumerate(lines):
        if line.startswith("Taiere"):
            start = idx + 2
            headers = ["Pos", "Material", "Grosime", "metri", "metri_per_qm",
                       "Lungime_per_Stck", "Zeit_min", "Pret_Std",
                       "Pret_metri", "Pret_qm", "Total"]
            rows, _ = parse_numeric_rows(lines, start, 9, headers)
            tables["Taiere"] = rows
            break

    # Bandarea marginilor
    for idx, line in enumerate(lines):
        if line.startswith("Bandarea marginilor"):
            start = idx + 2
            headers = ["Pos", "Material", "Grosime", "Inaltime", "Lungime_metri",
                       "per_Stck", "Bucata", "Zeit_min", "Pret_Std",
                       "Pret_metri", "AS_pct", "PE_metri", "Total"]
            rows, _ = parse_numeric_rows(lines, start, 11, headers)
            tables["Bandarea_marginilor"] = rows
            break

    # Costuri speciale
    for idx, line in enumerate(lines):
        if line.startswith("Costuri speciale"):
            start = idx + 2
            rows, _ = parse_costuri_speciale(lines, start)
            tables["Costuri_speciale"] = rows
            break

    # write CSVs
    for name, rows in tables.items():
        out_path = f"{out_dir}/{name}.csv"
        with open(out_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerows(rows)
        print(f"Wrote {out_path} ({len(rows)-1} data rows)")

    return tables


if __name__ == "__main__":
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else "input.pdf"
    out_dir = sys.argv[2] if len(sys.argv) > 2 else "."
    main(pdf_path, out_dir)
