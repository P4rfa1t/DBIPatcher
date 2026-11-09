"""Add the DBI 920+ save-data and ticket information labels.

These two screens are drawn from plain-ASCII .rodata labels ("Issuer          : ")
and, before the ctor_exact/pfxsfx ASCII gate was opened, no language could
translate them. Keys keep DBI's exact padding so the prefix match hits; the
translation keeps the same colon column by padding with spaces.

Run: python scripts/add_panel_labels.py [--dry-run]
"""
from __future__ import annotations

import csv
import json
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TR = ROOT / "translations"
DICT = ROOT / "data" / "dictionary.xlsx"
LANGS = list(json.loads((ROOT / "data" / "languages.json").read_text(encoding="utf-8")).keys())

# key -> {lang: value}. Keys are byte-exact copies of DBI's .rodata strings.
LABELS: dict[str, dict[str, str]] = {
    # ── Ticket information ────────────────────────────────────────────────
    "Signature       : ": {
        "ua": "Підпис          : ", "be": "Подпіс          : ", "de": "Signatur        : ",
        "en": "Signature       : ", "engb": "Signature       : ", "es": "Firma           : ",
        "es419": "Firma           : ", "et": "Allkiri         : ", "fr": "Signature       : ",
        "frca": "Signature       : ", "id": "Tanda tangan    : ", "it": "Firma           : ",
        "jp": "署名            : ", "kk": "Қолтаңба        : ", "kr": "서명            : ",
        "lt": "Parašas         : ", "lv": "Paraksts        : ", "nl": "Handtekening    : ",
        "pl": "Podpis          : ", "pt": "Assinatura      : ", "ptbr": "Assinatura      : ",
        "tr": "İmza            : ", "zhcn": "签名            : ", "zhtw": "簽章            : ",
    },
    "Issuer          : ": {
        "ua": "Видавець        : ", "be": "Выдавец         : ", "de": "Aussteller      : ",
        "en": "Issuer          : ", "engb": "Issuer          : ", "es": "Emisor          : ",
        "es419": "Emisor          : ", "et": "Väljaandja      : ", "fr": "Émetteur        : ",
        "frca": "Émetteur        : ", "id": "Penerbit        : ", "it": "Emittente       : ",
        "jp": "発行者          : ", "kk": "Беруші          : ", "kr": "발급자          : ",
        "lt": "Išdavėjas       : ", "lv": "Izdevējs        : ", "nl": "Uitgever        : ",
        "pl": "Wystawca        : ", "pt": "Emissor         : ", "ptbr": "Emissor         : ",
        "tr": "Veren           : ", "zhcn": "颁发者          : ", "zhtw": "簽發者          : ",
    },
    "Format version  : ": {
        "ua": "Версія формату  : ", "be": "Версія фармату  : ", "de": "Formatversion   : ",
        "en": "Format version  : ", "engb": "Format version  : ", "es": "Versión formato : ",
        "es419": "Versión formato : ", "et": "Vormingu versioon: ", "fr": "Version format  : ",
        "frca": "Version format  : ", "id": "Versi format    : ", "it": "Versione formato: ",
        "jp": "形式バージョン  : ", "kk": "Пішім нұсқасы   : ", "kr": "형식 버전       : ",
        "lt": "Formato versija : ", "lv": "Formāta versija : ", "nl": "Formaatversie   : ",
        "pl": "Wersja formatu : ", "pt": "Versão formato  : ", "ptbr": "Versão formato  : ",
        "tr": "Biçim sürümü    : ", "zhcn": "格式版本        : ", "zhtw": "格式版本        : ",
    },
    "Type            : ": {
        "ua": "Тип             : ", "be": "Тып             : ", "de": "Typ             : ",
        "en": "Type            : ", "engb": "Type            : ", "es": "Tipo            : ",
        "es419": "Tipo            : ", "et": "Tüüp            : ", "fr": "Type            : ",
        "frca": "Type            : ", "id": "Jenis           : ", "it": "Tipo            : ",
        "jp": "種類            : ", "kk": "Түрі            : ", "kr": "유형            : ",
        "lt": "Tipas           : ", "lv": "Tips            : ", "nl": "Type            : ",
        "pl": "Typ             : ", "pt": "Tipo            : ", "ptbr": "Tipo            : ",
        "tr": "Tür             : ", "zhcn": "类型            : ", "zhtw": "類型            : ",
    },
    "Version         : ": {
        "ua": "Версія          : ", "be": "Версія          : ", "de": "Version         : ",
        "en": "Version         : ", "engb": "Version         : ", "es": "Versión         : ",
        "es419": "Versión         : ", "et": "Versioon        : ", "fr": "Version         : ",
        "frca": "Version         : ", "id": "Versi           : ", "it": "Versione        : ",
        "jp": "バージョン      : ", "kk": "Нұсқасы         : ", "kr": "버전            : ",
        "lt": "Versija         : ", "lv": "Versija         : ", "nl": "Versie          : ",
        "pl": "Wersja          : ", "pt": "Versão          : ", "ptbr": "Versão          : ",
        "tr": "Sürüm           : ", "zhcn": "版本            : ", "zhtw": "版本            : ",
    },
    "License         : ": {
        "ua": "Ліцензія        : ", "be": "Ліцэнзія        : ", "de": "Lizenz          : ",
        "en": "License         : ", "engb": "Licence         : ", "es": "Licencia        : ",
        "es419": "Licencia        : ", "et": "Litsents        : ", "fr": "Licence         : ",
        "frca": "Licence         : ", "id": "Lisensi         : ", "it": "Licenza         : ",
        "jp": "ライセンス      : ", "kk": "Лицензия        : ", "kr": "라이선스        : ",
        "lt": "Licencija       : ", "lv": "Licence         : ", "nl": "Licentie        : ",
        "pl": "Licencja        : ", "pt": "Licença         : ", "ptbr": "Licença         : ",
        "tr": "Lisans          : ", "zhcn": "许可证          : ", "zhtw": "授權            : ",
    },
    "Key generation  : ": {
        "ua": "Покоління ключа : ", "be": "Пакаленне ключа : ", "de": "Schlüsselgen.   : ",
        "en": "Key generation  : ", "engb": "Key generation  : ", "es": "Gen. de clave   : ",
        "es419": "Gen. de clave   : ", "et": "Võtme põlvkond  : ", "fr": "Gén. de clé     : ",
        "frca": "Gén. de clé     : ", "id": "Gen. kunci      : ", "it": "Gen. chiave     : ",
        "jp": "キー世代        : ", "kk": "Кілт ұрпағы     : ", "kr": "키 세대         : ",
        "lt": "Raktų karta     : ", "lv": "Atslēgas paaudze: ", "nl": "Sleutelgen.     : ",
        "pl": "Gen. klucza     : ", "pt": "Ger. de chave   : ", "ptbr": "Ger. de chave   : ",
        "tr": "Anahtar nesli   : ", "zhcn": "密钥代号        : ", "zhtw": "金鑰世代        : ",
    },
    "Properties      : ": {
        "ua": "Властивості     : ", "be": "Уласцівасці     : ", "de": "Eigenschaften   : ",
        "en": "Properties      : ", "engb": "Properties      : ", "es": "Propiedades     : ",
        "es419": "Propiedades     : ", "et": "Omadused        : ", "fr": "Propriétés      : ",
        "frca": "Propriétés      : ", "id": "Properti        : ", "it": "Proprietà      : ",
        "jp": "プロパティ      : ", "kk": "Қасиеттері      : ", "kr": "속성            : ",
        "lt": "Savybės         : ", "lv": "Īpašības        : ", "nl": "Eigenschappen   : ",
        "pl": "Właściwości     : ", "pt": "Propriedades    : ", "ptbr": "Propriedades    : ",
        "tr": "Özellikler      : ", "zhcn": "属性            : ", "zhtw": "屬性            : ",
    },
    "Ticket ID       : ": {
        "ua": "ID квитка       : ", "be": "ID білета       : ", "de": "Ticket-ID       : ",
        "en": "Ticket ID       : ", "engb": "Ticket ID       : ", "es": "ID del ticket   : ",
        "es419": "ID del ticket   : ", "et": "Pileti ID       : ", "fr": "ID du ticket    : ",
        "frca": "ID du ticket    : ", "id": "ID tiket        : ", "it": "ID ticket       : ",
        "jp": "チケットID      : ", "kk": "Билет ID        : ", "kr": "티켓 ID         : ",
        "lt": "Bilieto ID      : ", "lv": "Biļetes ID      : ", "nl": "Ticket-ID       : ",
        "pl": "ID biletu       : ", "pt": "ID do ticket    : ", "ptbr": "ID do ticket    : ",
        "tr": "Bilet kimliği   : ", "zhcn": "票据ID          : ", "zhtw": "票券 ID         : ",
    },
    "Device ID       : ": {
        "ua": "ID пристрою     : ", "be": "ID прылады      : ", "de": "Geräte-ID       : ",
        "en": "Device ID       : ", "engb": "Device ID       : ", "es": "ID dispositivo  : ",
        "es419": "ID dispositivo  : ", "et": "Seadme ID       : ", "fr": "ID appareil     : ",
        "frca": "ID appareil     : ", "id": "ID perangkat    : ", "it": "ID dispositivo  : ",
        "jp": "デバイスID      : ", "kk": "Құрылғы ID      : ", "kr": "기기 ID         : ",
        "lt": "Įrenginio ID    : ", "lv": "Ierīces ID      : ", "nl": "Apparaat-ID     : ",
        "pl": "ID urządzenia   : ", "pt": "ID do dispositivo: ", "ptbr": "ID do dispositivo: ",
        "tr": "Cihaz kimliği   : ", "zhcn": "设备ID          : ", "zhtw": "裝置 ID         : ",
    },
    "RightsId        : ": {
        "ua": "RightsId        : ", "be": "RightsId        : ", "de": "RightsId        : ",
        "en": "RightsId        : ", "engb": "RightsId        : ", "es": "RightsId        : ",
        "es419": "RightsId        : ", "et": "RightsId        : ", "fr": "RightsId        : ",
        "frca": "RightsId        : ", "id": "RightsId        : ", "it": "RightsId        : ",
        "jp": "RightsId        : ", "kk": "RightsId        : ", "kr": "RightsId        : ",
        "lt": "RightsId        : ", "lv": "RightsId        : ", "nl": "RightsId        : ",
        "pl": "RightsId        : ", "pt": "RightsId        : ", "ptbr": "RightsId        : ",
        "tr": "RightsId        : ", "zhcn": "权利ID          : ", "zhtw": "權利 ID         : ",
    },
    "Enc. title key  : ": {
        "ua": "Шифр. ключ титлу: ", "be": "Шыфр. ключ тытлу: ", "de": "Verschl. Titelkey: ",
        "en": "Enc. title key  : ", "engb": "Enc. title key  : ", "es": "Clave cifrada   : ",
        "es419": "Clave cifrada   : ", "et": "Krüpt. pealkirja võti: ", "fr": "Clé chiffrée    : ",
        "frca": "Clé chiffrée    : ", "id": "Kunci terenkripsi: ", "it": "Chiave cifrata  : ",
        "jp": "暗号化タイトルキー: ", "kk": "Шифр. атау кілті: ", "kr": "암호화 제목 키  : ",
        "lt": "Šifr. pavad. raktas: ", "lv": "Šifr. virsraksta atslēga: ", "nl": "Versl. titelsleutel: ",
        "pl": "Zasz. klucz tytułu: ", "pt": "Chave cifrada   : ", "ptbr": "Chave cifrada   : ",
        "tr": "Şifr. başlık anahtarı: ", "zhcn": "加密标题密钥    : ", "zhtw": "加密標題金鑰    : ",
    },
    "Dec. title key  : ": {
        "ua": "Розш. ключ титлу: ", "be": "Расш. ключ тытлу: ", "de": "Entschl. Titelkey: ",
        "en": "Dec. title key  : ", "engb": "Dec. title key  : ", "es": "Clave descifrada: ",
        "es419": "Clave descifrada: ", "et": "Dekrüpt. pealkirja võti: ", "fr": "Clé déchiffrée  : ",
        "frca": "Clé déchiffrée  : ", "id": "Kunci didekripsi: ", "it": "Chiave decifrata: ",
        "jp": "復号タイトルキー: ", "kk": "Дешифр. атау кілті: ", "kr": "복호화 제목 키  : ",
        "lt": "Dešifr. pavad. raktas: ", "lv": "Atšifr. virsraksta atslēga: ", "nl": "Ontsleutelde titelsleutel: ",
        "pl": "Odsz. klucz tytułu: ", "pt": "Chave decifrada : ", "ptbr": "Chave decifrada : ",
        "tr": "Deşifre başlık anahtarı: ", "zhcn": "解密标题密钥    : ", "zhtw": "解密標題金鑰    : ",
    },
    # ── Save-data information ─────────────────────────────────────────────
    "Space  : ": {
        "ua": "Місце  : ", "be": "Месца  : ", "de": "Speicher : ", "en": "Space  : ",
        "engb": "Space  : ", "es": "Espacio: ", "es419": "Espacio: ", "et": "Ruum   : ",
        "fr": "Espace : ", "frca": "Espace : ", "id": "Ruang  : ", "it": "Spazio : ",
        "jp": "容量   : ", "kk": "Орын   : ", "kr": "공간   : ", "lt": "Vieta  : ",
        "lv": "Vieta  : ", "nl": "Ruimte : ", "pl": "Miejsce: ", "pt": "Espaço : ",
        "ptbr": "Espaço : ", "tr": "Alan   : ", "zhcn": "空间   : ", "zhtw": "空間   : ",
    },
    "Type   : ": {
        "ua": "Тип    : ", "be": "Тып    : ", "de": "Typ    : ", "en": "Type   : ",
        "engb": "Type   : ", "es": "Tipo   : ", "es419": "Tipo   : ", "et": "Tüüp   : ",
        "fr": "Type   : ", "frca": "Type   : ", "id": "Jenis  : ", "it": "Tipo   : ",
        "jp": "種類   : ", "kk": "Түрі   : ", "kr": "유형   : ", "lt": "Tipas  : ",
        "lv": "Tips   : ", "nl": "Type   : ", "pl": "Typ    : ", "pt": "Tipo   : ",
        "ptbr": "Tipo   : ", "tr": "Tür    : ", "zhcn": "类型   : ", "zhtw": "類型   : ",
    },
    "User   : ": {
        "ua": "Користувач: ", "be": "Карыстальнік: ", "de": "Nutzer : ", "en": "User   : ",
        "engb": "User   : ", "es": "Usuario: ", "es419": "Usuario: ", "et": "Kasutaja: ",
        "fr": "User   : ", "frca": "User   : ", "id": "Pengguna: ", "it": "Utente : ",
        "jp": "ユーザー: ", "kk": "Пайдаланушы: ", "kr": "사용자 : ", "lt": "Vartotojas: ",
        "lv": "Lietotājs: ", "nl": "Gebruiker: ", "pl": "Użytkownik: ", "pt": "Usuário: ",
        "ptbr": "Usuário: ", "tr": "Kullanıcı: ", "zhcn": "用户   : ", "zhtw": "使用者 : ",
    },
    "Uid    : ": {
        "ua": "Uid    : ", "be": "Uid    : ", "de": "Uid    : ", "en": "Uid    : ",
        "engb": "Uid    : ", "es": "Uid    : ", "es419": "Uid    : ", "et": "Uid    : ",
        "fr": "Uid    : ", "frca": "Uid    : ", "id": "Uid    : ", "it": "Uid    : ",
        "jp": "Uid    : ", "kk": "Uid    : ", "kr": "Uid    : ", "lt": "Uid    : ",
        "lv": "Uid    : ", "nl": "Uid    : ", "pl": "Uid    : ", "pt": "Uid    : ",
        "ptbr": "Uid    : ", "tr": "Uid    : ", "zhcn": "Uid    : ", "zhtw": "Uid    : ",
    },
    "Index  : ": {
        "ua": "Індекс : ", "be": "Індэкс : ", "de": "Index  : ", "en": "Index  : ",
        "engb": "Index  : ", "es": "Índice : ", "es419": "Índice : ", "et": "Indeks : ",
        "fr": "Index  : ", "frca": "Index  : ", "id": "Indeks : ", "it": "Indice : ",
        "jp": "索引   : ", "kk": "Индекс : ", "kr": "색인   : ", "lt": "Indeksas: ",
        "lv": "Indekss: ", "nl": "Index  : ", "pl": "Indeks : ", "pt": "Índice : ",
        "ptbr": "Índice : ", "tr": "Dizin  : ", "zhcn": "索引   : ", "zhtw": "索引   : ",
    },
}


def main() -> None:
    dry = "--dry-run" in sys.argv
    missing = {k: sorted(set(LANGS) - set(v)) for k, v in LABELS.items()}
    missing = {k: v for k, v in missing.items() if v}
    if missing:
        sys.exit(f"missing languages: {missing}")

    wb = openpyxl.load_workbook(DICT)
    ws = wb["Translations"]
    col = {h: i + 1 for i, h in enumerate(c.value for c in ws[1]) if h}

    existing = set()
    for r in range(2, ws.max_row + 1):
        o = ws.cell(r, col["Original"]).value
        if o is not None:
            existing.add(str(o))

    added = 0
    for key, per_lang in LABELS.items():
        if key in existing:
            print(f"skip (already present): {key!r}")
            continue
        if not dry:
            r = ws.max_row + 1
            ws.cell(r, col["Original"], key)
            for lang, text in per_lang.items():
                ws.cell(r, col[lang], text)
        added += 1
    if not dry:
        wb.save(DICT)
    print(f"{'would add' if dry else 'added'} {added} label keys; workbook rows now {ws.max_row - 1}")


if __name__ == "__main__":
    main()
