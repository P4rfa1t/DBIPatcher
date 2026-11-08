"""Fix two real gaps found by the audit:

1. zhcn / zhtw: the Shadok-parody "publisher" line (Жак Руксель) falls outside
   shadok.json's mapping, so it was never localised. en/ua carry real text there.
2. zhtw: the "Великий колдун сказал" line of the same embedded paragraph is a
   bare space while the rest of the paragraph is translated.

Applies to dictionary.xlsx (source of truth for export) and the CSVs, then
rebuilds the affected binaries.
"""
from __future__ import annotations

import csv
import subprocess
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
DICT = ROOT / "data" / "dictionary.xlsx"

FIXES = {
    "  Жак Руксель. Великолепие навыворот.": {
        "zhcn": "  地堡，镜中世界国营出版社.",
    },
    "    Великий колдун сказал, что ракеты": {
        "zhtw": "偉大的巫師宣稱，火箭之所以",
    },
}


def main() -> None:
    wb = openpyxl.load_workbook(DICT)
    ws = wb["Translations"]
    headers = [c.value for c in ws[1]]
    col = {h: i + 1 for i, h in enumerate(headers) if h}

    rows = {}
    for r in range(2, ws.max_row + 1):
        o = ws.cell(r, col["Original"]).value
        if o is not None:
            rows.setdefault(str(o), []).append(r)

    changed = 0
    for orig, per_lang in FIXES.items():
        hits = rows.get(orig, [])
        if len(hits) != 1:
            raise SystemExit(f"expected one workbook row for {orig!r}, found {hits}")
        r = hits[0]
        for lang, text in per_lang.items():
            cur = ws.cell(r, col[lang]).value
            if cur and str(cur).strip() and str(cur) != text:
                print(f"  skip {lang} row {r}: already has {cur!r}")
                continue
            ws.cell(r, col[lang], text)
            changed += 1
            print(f"  set {lang} row {r}: {orig[:40]!r} -> {text!r}")
    wb.save(DICT)
    print(f"dictionary.xlsx: {changed} cells updated")

    subprocess.run([sys.executable, "-m", "src.main", "export"], cwd=ROOT, check=True)


if __name__ == "__main__":
    main()
