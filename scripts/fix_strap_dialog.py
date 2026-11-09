"""Fix the meaning of "Не показывать диалог про верёвочку".

"верёвочка" (little rope) is DBI's nickname for the Joy-Con wrist strap: the
toggle blocks nn::hid::ShowControllerStrapGuide, the system dialog that asks the
player to attach the strap. The previous wording ("reboot dialog") came from a
blurry screenshot and was wrong in every language.
"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
DICT = ROOT / "data" / "dictionary.xlsx"
KEY = "Не показывать диалог про верёвочку"

CORRECT = {
    "ua": "Не показувати діалог про ремінець", "be": "Не паказваць дыялог пра раменьчык",
    "de": "Handschlaufen-Dialog nicht anzeigen", "en": "Do not show the wrist strap dialog",
    "engb": "Do not show the wrist strap dialog", "es": "No mostrar el diálogo de la correa",
    "es419": "No mostrar el diálogo de la correa", "et": "Ära näita kinnitusrihma dialoogi",
    "fr": "Ne pas afficher la boîte de la dragonne",
    "frca": "Ne pas afficher la boîte de la dragonne",
    "id": "Jangan tampilkan dialog tali pengikat", "it": "Non mostrare la finestra della cinghia",
    "jp": "リストストラップのダイアログを表示しない", "kk": "Білезік бауы диалогын көрсетпеу",
    "kr": "손목 스트랩 대화상자 표시 안 함", "lt": "Nerodyti riešo dirželio dialogo",
    "lv": "Nerādīt aproces dialoglodziņu", "nl": "Polsbandje-dialoog niet tonen",
    "pl": "Nie pokazuj okna paska na rękę", "pt": "Não mostrar o diálogo da alça",
    "ptbr": "Não mostrar o diálogo da alça", "tr": "Bileklik kayışı iletişim kutusunu gösterme",
    "zhcn": "不显示腕带提示对话框", "zhtw": "不顯示腕帶提示對話框",
}


def main() -> None:
    langs = json.loads((ROOT / "data" / "languages.json").read_text(encoding="utf-8"))
    missing = sorted(set(langs) - set(CORRECT))
    if missing:
        sys.exit(f"missing languages: {missing}")

    wb = openpyxl.load_workbook(DICT)
    ws = wb["Translations"]
    col = {h: i + 1 for i, h in enumerate(c.value for c in ws[1]) if h}

    hit = None
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, col["Original"]).value == KEY:
            hit = r
            break
    if hit is None:
        sys.exit(f"row not found: {KEY!r}")

    for lang, text in CORRECT.items():
        ws.cell(hit, col[lang], text)
    wb.save(DICT)
    print(f"row {hit}: rewrote {len(CORRECT)} translations for {KEY!r}")

    subprocess.run([sys.executable, "-m", "src.main", "export"], cwd=ROOT, check=True)


if __name__ == "__main__":
    main()
