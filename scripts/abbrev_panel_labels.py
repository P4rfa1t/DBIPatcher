"""Shorten panel-label translations that cannot fit their DBI cell column.

DBI pads these labels to a fixed column; a word longer than the column pushes the
value out of line. Abbreviate with a trailing dot, which is the convention the
template already uses ("Enc. title key", "Dec. title key").
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
DICT = ROOT / "data" / "dictionary.xlsx"

# label -> {lang: text that must fit}. Only entries the column has no room for.
ABBREV = {
    "Dec. title key  : ": {"tr": "Deşif. baş. an.", "et": "Dekrüpt. pealk. v.",
                           "kk": "Дешифр. атау кілті"},
    "Enc. title key  : ": {"tr": "Şifr. baş. an.", "et": "Krüpt. pealk. v.",
                           "kk": "Шифр. атау кілті"},
    "User   : ": {"et": "Kasut."},
}


def main() -> None:
    wb = openpyxl.load_workbook(DICT)
    ws = wb["Translations"]
    col = {h: i + 1 for i, h in enumerate(c.value for c in ws[1]) if h}

    rows: dict[str, int] = {}
    for r in range(2, ws.max_row + 1):
        o = ws.cell(r, col["Original"]).value
        if o is not None:
            rows.setdefault(str(o), r)

    changed = 0
    for orig, per_lang in ABBREV.items():
        r = rows.get(orig)
        if r is None:
            continue
        width = orig.index(":")
        for lang, text in per_lang.items():
            if len(text) > width:
                print(f"  [WARN] {lang}: {text!r} is {len(text)} > {width}; skipped")
                continue
            ws.cell(r, col[lang], text + " " * (width - len(text)) + ": ")
            changed += 1
            print(f"  {lang} row {r}: {orig.strip()!r} -> {text!r}")
    wb.save(DICT)
    print(f"{changed} cells rewritten")
    if "--export" in sys.argv:
        subprocess.run([sys.executable, "-m", "src.main", "export"], cwd=ROOT, check=True)


if __name__ == "__main__":
    main()
