"""Pad the new panel-label translations so their colon lands in the same cell
column as the source. DBI draws every glyph in a fixed cell, so the label column
must line up or the value column staggers.
"""
from __future__ import annotations

import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
DICT = ROOT / "data" / "dictionary.xlsx"

PANEL = [
    "Signature       : ", "Issuer          : ", "Format version  : ", "Type            : ",
    "Version         : ", "License         : ", "Key generation  : ", "Properties      : ",
    "Ticket ID       : ", "Device ID       : ", "RightsId        : ", "Enc. title key  : ",
    "Dec. title key  : ", "Space  : ", "Type   : ", "User   : ", "Uid    : ", "Index  : ",
]


def main() -> None:
    wb = openpyxl.load_workbook(DICT)
    ws = wb["Translations"]
    col = {h: i + 1 for i, h in enumerate(c.value for c in ws[1]) if h}
    langs = [h for h in col if h != "Original" and h != "ru"]

    fixed = 0
    for r in range(2, ws.max_row + 1):
        orig = ws.cell(r, col["Original"]).value
        if orig not in PANEL:
            continue
        width = str(orig).index(":")
        for lang in langs:
            val = ws.cell(r, col[lang]).value
            if not val:
                continue
            text = str(val)
            if ":" not in text:
                continue
            label = text[:text.index(":")].rstrip()
            if len(label) > width:
                print(f"  [WARN] {lang} row {r}: label wider than {width}: {label!r}")
                continue
            new = label + " " * (width - len(label)) + text[text.index(":"):]
            if new != text:
                ws.cell(r, col[lang], new)
                fixed += 1
    wb.save(DICT)
    print(f"padded {fixed} cells")


if __name__ == "__main__":
    if "--apply" in sys.argv:
        main()
