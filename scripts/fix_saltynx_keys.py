"""Replace the four SaltyNX labels whose Russian text was misread from a blurry
screenshot. The dialog draws 'ё' and uses different wording than the keys added
previously, so those four rows never matched.
"""
from __future__ import annotations

import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
DICT = ROOT / "data" / "dictionary.xlsx"

RENAME = {
    "Форсировать режим полета": "Форсировать режим полёта",
    "Не показывать диалоги с ошибками": "Не показывать диалоги с ошибкой",
    "Не показывать диалог про перезагрузку": "Не показывать диалог про верёвочку",
    "Блокировать изменение частот CPU": "Блокировать изменение частот APU",
}


def main() -> None:
    wb = openpyxl.load_workbook(DICT)
    ws = wb["Translations"]
    hdr = [c.value for c in ws[1]]
    col = {h: i + 1 for i, h in enumerate(hdr) if h}

    rows: dict[str, list[int]] = {}
    for r in range(2, ws.max_row + 1):
        o = ws.cell(r, col["Original"]).value
        if o is not None:
            rows.setdefault(str(o), []).append(r)

    for old, new in RENAME.items():
        hits = rows.get(old, [])
        if len(hits) != 1:
            sys.exit(f"expected one row for {old!r}, got {hits}")
        if new in rows:
            sys.exit(f"{new!r} already exists at rows {rows[new]}")
        ws.cell(hits[0], col["Original"], new)
        print(f"row {hits[0]}: {old!r} -> {new!r}")

    wb.save(DICT)
    print("dictionary.xlsx updated")


if __name__ == "__main__":
    main()
