"""Guard the SaltyNX patch-menu strings and catch silent untranslated fallbacks.

The SaltyNX patch menu (title + twelve toggle labels) is drawn from DBI's
decrypted string store, so it was invisible to every earlier audit: the keys
simply did not exist in the dictionary and all 23 languages fell back to
Russian. This test pins the keys, checks each language has them, and rejects a
translation that silently equals the English fallback or is left blank outside
the Shadok parody block.
"""
from __future__ import annotations

import csv
import json
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TRANSLATIONS = ROOT / "translations"
LANGS = json.loads((ROOT / "data" / "languages.json").read_text(encoding="utf-8"))

SALTYNX_KEYS = (
    "Патчи SaltyNX",
    "Форсировать режим полёта",
    "Игнорировать ответ от сервера",
    "Считать аккаунт линкованным",
    "Не отправлять запросы на сервер",
    "Блокировать все сетевые действия",
    "Не показывать диалоги с ошибкой",
    "Убрать копирайт со скриншотов",
    "Блокировать отключение видеозаписи",
    "Блокировать отключение скриншотов",
    "Не показывать диалог про верёвочку",
    "Не останавливать игру в фоне",
    "Блокировать изменение частот APU",
)

# DBI 920+ save-data and ticket information panels: plain-ASCII labels that the
# runtime used to skip (see the ASCII gate in vendor/dbi_translate/runtime.py).
PANEL_LABELS = (
    "Signature       : ",
    "Issuer          : ",
    "Format version  : ",
    "Type            : ",
    "Version         : ",
    "License         : ",
    "Key generation  : ",
    "Properties      : ",
    "Ticket ID       : ",
    "Device ID       : ",
    "RightsId        : ",
    "Enc. title key  : ",
    "Dec. title key  : ",
    "Space  : ",
    "Type   : ",
    "User   : ",
    "Uid    : ",
    "Index  : ",
)

ASCII_IDENTICAL_OK = frozenset({"Х:{}, Т:{}, РЦ:{}", "МБ/сек)", "ru", "original"})


def load(lang: str) -> dict[str, str]:
    with (TRANSLATIONS / f"{lang}.csv").open(encoding="utf-8", newline="") as f:
        return {row["original"]: row["translation"] for row in csv.DictReader(f)}


class SaltyNXPatchMenuTests(unittest.TestCase):
    def test_every_language_has_the_patch_menu_labels(self) -> None:
        missing = []
        for lang in LANGS:
            table = load(lang)
            for key in SALTYNX_KEYS:
                if not table.get(key, "").strip():
                    missing.append(f"{lang}: {key!r}")
        self.assertEqual(missing, [], f"missing SaltyNX labels: {missing[:10]}")

    def test_patch_menu_labels_are_not_left_in_russian(self) -> None:
        """Non-Cyrillic-target languages must not repeat the Russian source."""
        for lang in ("en", "engb", "de", "fr", "frca", "es", "es419", "it", "pt",
                     "ptbr", "nl", "pl", "tr", "id", "et", "lt", "lv", "jp", "kr",
                     "zhcn", "zhtw"):
            table = load(lang)
            same = [k for k in SALTYNX_KEYS if table.get(k) == k]
            self.assertEqual(same, [], f"{lang} left these untranslated: {same}")

    def test_chinese_labels_render_in_han(self) -> None:
        for lang in ("zhcn", "zhtw"):
            table = load(lang)
            for key in SALTYNX_KEYS:
                value = table[key]
                self.assertTrue(
                    any("一" <= c <= "鿿" for c in value),
                    f"{lang}: {key!r} -> {value!r} has no Han characters",
                )

    def test_panel_labels_present_in_every_language(self) -> None:
        missing = []
        for lang in LANGS:
            table = load(lang)
            for key in PANEL_LABELS:
                if not table.get(key, "").strip():
                    missing.append(f"{lang}: {key!r}")
        self.assertEqual(missing, [], f"missing panel labels: {missing[:10]}")

    def test_panel_labels_keep_the_colon_column(self) -> None:
        """DBI pads these labels to a fixed cell column; a translation that shifts the colon
        staggers the value column. Only zhcn/zhtw are held to it: the CJK words are short
        enough to pad, while several Latin/Cyrillic words cannot fit and are left to their
        own translators."""
        bad = []
        for lang in ("zhcn", "zhtw"):
            table = load(lang)
            for key in PANEL_LABELS:
                value = table.get(key, "")
                if not value or ":" not in value:
                    continue
                source_col, value_col = key.index(":"), value.index(":")
                if value_col > source_col:
                    bad.append(f"{lang}: {key!r} -> {value!r} (colon {value_col} > {source_col})")
        self.assertEqual(bad, [], f"colon overflowed: {bad[:10]}")


class UntranslatedFallbackTests(unittest.TestCase):
    """A blank or English-identical cell means export silently fell back."""

    # Languages that must never contain Latin-script English in place of Han/kana/
    # hangul/Cyrillic text: an identical English copy there is always a fallback.
    SCRIPT_LANGS = ("jp", "kr", "zhcn", "zhtw", "ua", "be", "kk")

    def test_parody_paragraph_lines_are_filled(self) -> None:
        """The embedded Shadok paragraph has lines outside shadok.json's mapping.

        Line 1124 (the "publisher" line) and the "Великий колдун" line read as
        Russian in de/fr/jp/... when blank, because export only falls back to EN.
        """
        required = {
            "zhcn": ("  Жак Руксель. Великолепие навыворот.",),
            "zhtw": ("  Жак Руксель. Великолепие навыворот.",
                     "    Великий колдун сказал, что ракеты"),
        }
        for lang, keys in required.items():
            table = load(lang)
            for key in keys:
                self.assertTrue(table.get(key, "").strip(),
                                f"{lang}: {key!r} is still blank")

    def test_script_languages_are_not_silent_english_copies(self) -> None:
        english = load("en")
        copied = []
        for lang in self.SCRIPT_LANGS:
            for key, value in load(lang).items():
                if key in ASCII_IDENTICAL_OK:
                    continue
                if value == english.get(key) and value != key:
                    copied.append(f"{lang}: {key!r} -> {value!r}")
        self.assertEqual(copied, [], f"English fallbacks: {copied[:10]}")


if __name__ == "__main__":
    unittest.main()
