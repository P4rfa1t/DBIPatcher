"""Replace the two верёвочку rows with the correct перезагрузку key.

The tenth SaltyNX toggle reads "Не показывать диалог про перезагрузку" (confirmed
on the console by the reporter); the верёвочку variant was a misreading of a
blurry screenshot and must not ship.
"""
from __future__ import annotations

import subprocess
import sys
import unicodedata
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
DICT = ROOT / "data" / "dictionary.xlsx"

WRONG = ["Не показывать диалог про верёвочку"]
RIGHT = {
    "ua": "Не показувати діалог про перезавантаження", "be": "Не паказваць дыялог пра перазагрузку",
    "de": "Neustart-Dialog nicht anzeigen", "en": "Do not show reboot dialog",
    "engb": "Do not show reboot dialog", "es": "No mostrar diálogo de reinicio",
    "es419": "No mostrar diálogo de reinicio", "et": "Ära näita taaskäivituse dialoogi",
    "fr": "Ne pas afficher la boîte de redémarrage", "frca": "Ne pas afficher la boîte de redémarrage",
    "id": "Jangan tampilkan dialog mulai ulang", "it": "Non mostrare finestra di riavvio",
    "jp": "再起動ダイアログを表示しない", "kk": "Қайта жүктеу диалогын көрсетпеу",
    "kr": "재시작 대화상자 표시 안 함", "lt": "Nerodyti perkrovimo dialogo",
    "lv": "Nerādīt pārstartēšanas dialoglodziņu", "nl": "Herstartdialoog niet tonen",
    "pl": "Nie pokazuj okna ponownego uruchomienia", "pt": "Não mostrar diálogo de reinício",
    "ptbr": "Não mostrar diálogo de reinicialização", "tr": "Yeniden başlatma iletişim kutusunu gösterme",
    "zhcn": "不显示重启对话框", "zhtw": "不顯示重新啟動對話框",
}
KEY = "Не показывать диалог про перезагрузку"


def main() -> None:
    wb = openpyxl.load_workbook(DICT)
    ws = wb["Translations"]
    col = {h: i + 1 for i, h in enumerate(c.value for c in ws[1]) if h}

    drop = {unicodedata.normalize(f, w) for w in WRONG for f in ("NFC", "NFD")}
    removed = 0
    for r in range(ws.max_row, 1, -1):
        if ws.cell(r, col["Original"]).value in drop:
            print(f"  remove row {r}: {ws.cell(r, col['Original']).value!r}")
            ws.delete_rows(r)
            removed += 1

    forms = [unicodedata.normalize("NFC", KEY)]
    nfd = unicodedata.normalize("NFD", KEY)
    if nfd != forms[0]:
        forms.append(nfd)

    present = {ws.cell(r, col["Original"]).value for r in range(2, ws.max_row + 1)}
    added = 0
    for form in forms:
        if form in present:
            continue
        r = ws.max_row + 1
        ws.cell(r, col["Original"], form)
        for lang, text in RIGHT.items():
            ws.cell(r, col[lang], text)
        added += 1

    wb.save(DICT)
    print(f"removed {removed}, added {added}, rows now {ws.max_row - 1}")

    subprocess.run([sys.executable, "-m", "src.main", "export"], cwd=ROOT, check=True)


if __name__ == "__main__":
    main()
