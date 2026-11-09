"""Fill the SaltyNX patch-menu rows in every language.

Scope is exactly the SaltyNX dialog: its title plus the twelve toggles. Nothing
else is translated, and no other screen is touched. Keys keep DBI's wording,
including "полёта" and "про верёвочку"; ё-bearing rows are also added in NFD so
the byte-exact matcher hits whichever composition DBI's store holds.
"""
from __future__ import annotations

import csv
import json
import subprocess
import sys
import unicodedata
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
DICT = ROOT / "data" / "dictionary.xlsx"
TR = ROOT / "translations"

# key -> {lang: text}
T: dict[str, dict[str, str]] = {
    "Патчи SaltyNX": {
        "ua": "Патчі SaltyNX", "be": "Патчы SaltyNX", "de": "SaltyNX-Patches",
        "en": "SaltyNX patches", "engb": "SaltyNX patches", "es": "Parches de SaltyNX",
        "es419": "Parches de SaltyNX", "et": "SaltyNX plaastrid", "fr": "Correctifs SaltyNX",
        "frca": "Correctifs SaltyNX", "id": "Patch SaltyNX", "it": "Patch SaltyNX",
        "jp": "SaltyNX パッチ", "kk": "SaltyNX патчтары", "kr": "SaltyNX 패치",
        "lt": "SaltyNX pataisos", "lv": "SaltyNX ielāpi", "nl": "SaltyNX-patches",
        "pl": "Łatki SaltyNX", "pt": "Patches do SaltyNX", "ptbr": "Patches do SaltyNX",
        "tr": "SaltyNX yamaları", "zhcn": "SaltyNX 补丁", "zhtw": "SaltyNX 補丁",
    },
    "Форсировать режим полёта": {
        "ua": "Форсувати режим польоту", "be": "Фарсіраваць рэжым палёту",
        "de": "Flugmodus erzwingen", "en": "Force flight mode", "engb": "Force flight mode",
        "es": "Forzar modo avión", "es419": "Forzar modo avión", "et": "Sunni lennurežiimi",
        "fr": "Forcer le mode avion", "frca": "Forcer le mode avion",
        "id": "Paksa mode pesawat", "it": "Forza modalità aereo", "jp": "機内モードを強制",
        "kk": "Ұшу режимін мәжбүрлеу", "kr": "비행기 모드 강제",
        "lt": "Priverstinis skrydžio režimas", "lv": "Piespiedu lidojuma režīms",
        "nl": "Vliegtuigmodus forceren", "pl": "Wymuś tryb samolotowy",
        "pt": "Forçar modo de voo", "ptbr": "Forçar modo avião", "tr": "Uçuş modunu zorla",
        "zhcn": "强制飞行模式", "zhtw": "強制飛航模式",
    },
    "Игнорировать ответ от сервера": {
        "ua": "Ігнорувати відповідь від сервера", "be": "Ігнараваць адказ ад сервера",
        "de": "Serverantwort ignorieren", "en": "Ignore server response",
        "engb": "Ignore server response", "es": "Ignorar respuesta del servidor",
        "es419": "Ignorar respuesta del servidor", "et": "Eira serveri vastust",
        "fr": "Ignorer la réponse du serveur", "frca": "Ignorer la réponse du serveur",
        "id": "Abaikan respons server", "it": "Ignora risposta del server",
        "jp": "サーバー応答を無視", "kk": "Сервер жауабын елемеу", "kr": "서버 응답 무시",
        "lt": "Nepaisyti serverio atsakymo", "lv": "Ignorēt servera atbildi",
        "nl": "Serverantwoord negeren", "pl": "Ignoruj odpowiedź serwera",
        "pt": "Ignorar resposta do servidor", "ptbr": "Ignorar resposta do servidor",
        "tr": "Sunucu yanıtını yoksay", "zhcn": "忽略服务器响应", "zhtw": "忽略伺服器回應",
    },
    "Считать аккаунт линкованным": {
        "ua": "Вважати акаунт прив'язаним", "be": "Лічыць акаўнт прывязаным",
        "de": "Konto als verknüpft betrachten", "en": "Treat account as linked",
        "engb": "Treat account as linked", "es": "Considerar cuenta vinculada",
        "es419": "Considerar cuenta vinculada", "et": "Käsitle kontot seotuna",
        "fr": "Considérer le compte comme lié", "frca": "Considérer le compte comme lié",
        "id": "Anggap akun tertaut", "it": "Considera account collegato",
        "jp": "アカウント連携済みとみなす", "kk": "Аккаунтты байланысқан деп есептеу",
        "kr": "계정 연동됨으로 처리", "lt": "Laikyti paskyrą susietą",
        "lv": "Uzskatīt kontu par saistītu", "nl": "Account als gekoppeld beschouwen",
        "pl": "Traktuj konto jako powiązane", "pt": "Considerar conta vinculada",
        "ptbr": "Considerar conta vinculada", "tr": "Hesabı bağlı say",
        "zhcn": "视为已关联账户", "zhtw": "視為已連結帳號",
    },
    "Не отправлять запросы на сервер": {
        "ua": "Не надсилати запити на сервер", "be": "Не адпраўляць запыты на сервер",
        "de": "Keine Anfragen an den Server senden", "en": "Do not send requests to the server",
        "engb": "Do not send requests to the server", "es": "No enviar solicitudes al servidor",
        "es419": "No enviar solicitudes al servidor", "et": "Ära saada serverile päringuid",
        "fr": "Ne pas envoyer de requêtes au serveur", "frca": "Ne pas envoyer de requêtes au serveur",
        "id": "Jangan kirim permintaan ke server", "it": "Non inviare richieste al server",
        "jp": "サーバーへリクエストを送信しない", "kk": "Серверге сұрау жібермеу",
        "kr": "서버로 요청을 보내지 않음", "lt": "Nesiųsti užklausų serveriui",
        "lv": "Nesūtīt pieprasījumus serverim", "nl": "Geen verzoeken naar de server sturen",
        "pl": "Nie wysyłaj zapytań do serwera", "pt": "Não enviar pedidos ao servidor",
        "ptbr": "Não enviar solicitações ao servidor", "tr": "Sunucuya istek gönderme",
        "zhcn": "不向服务器发送请求", "zhtw": "不向伺服器傳送請求",
    },
    "Блокировать все сетевые действия": {
        "ua": "Блокувати всі мережеві дії", "be": "Блакаваць усе сеткавыя дзеянні",
        "de": "Alle Netzwerkaktionen blockieren", "en": "Block all network activity",
        "engb": "Block all network activity", "es": "Bloquear toda actividad de red",
        "es419": "Bloquear toda actividad de red", "et": "Blokeeri kogu võrgutegevus",
        "fr": "Bloquer toute activité réseau", "frca": "Bloquer toute activité réseau",
        "id": "Blokir semua aktivitas jaringan", "it": "Blocca tutte le attività di rete",
        "jp": "すべてのネットワーク動作をブロック", "kk": "Барлық желілік әрекеттерді блоктау",
        "kr": "모든 네트워크 동작 차단", "lt": "Blokuoti visą tinklo veiklą",
        "lv": "Bloķēt visas tīkla darbības", "nl": "Alle netwerkactiviteit blokkeren",
        "pl": "Blokuj całą aktywność sieciową", "pt": "Bloquear toda atividade de rede",
        "ptbr": "Bloquear toda atividade de rede", "tr": "Tüm ağ etkinliğini engelle",
        "zhcn": "阻止所有网络活动", "zhtw": "封鎖所有網路活動",
    },
    "Не показывать диалоги с ошибкой": {
        "ua": "Не показувати діалоги з помилкою", "be": "Не паказваць дыялогі з памылкай",
        "de": "Fehlerdialoge nicht anzeigen", "en": "Do not show error dialogs",
        "engb": "Do not show error dialogs", "es": "No mostrar diálogos de error",
        "es419": "No mostrar diálogos de error", "et": "Ära näita veateateid",
        "fr": "Ne pas afficher les boîtes d'erreur", "frca": "Ne pas afficher les boîtes d'erreur",
        "id": "Jangan tampilkan dialog galat", "it": "Non mostrare finestre di errore",
        "jp": "エラーダイアログを表示しない", "kk": "Қате хабарламаларын көрсетпеу",
        "kr": "오류 대화상자 표시 안 함", "lt": "Nerodyti klaidų dialogų",
        "lv": "Nerādīt kļūdu dialoglodziņus", "nl": "Foutdialogen niet tonen",
        "pl": "Nie pokazuj okien błędów", "pt": "Não mostrar diálogos de erro",
        "ptbr": "Não mostrar diálogos de erro", "tr": "Hata iletişim kutularını gösterme",
        "zhcn": "不显示错误对话框", "zhtw": "不顯示錯誤對話框",
    },
    "Убрать копирайт со скриншотов": {
        "ua": "Прибрати копірайт зі скриншотів", "be": "Прыбраць капірайт са скрыншотаў",
        "de": "Copyright von Screenshots entfernen", "en": "Remove copyright from screenshots",
        "engb": "Remove copyright from screenshots", "es": "Quitar copyright de las capturas",
        "es419": "Quitar copyright de las capturas", "et": "Eemalda ekraanipiltidelt autoriõigus",
        "fr": "Retirer le copyright des captures", "frca": "Retirer le copyright des captures",
        "id": "Hapus hak cipta dari tangkapan layar", "it": "Rimuovi copyright dagli screenshot",
        "jp": "スクリーンショットの著作権表示を削除", "kk": "Скриншоттардан копирайтты алып тастау",
        "kr": "스크린샷에서 저작권 표시 제거", "lt": "Pašalinti autorių teises iš ekrano kopijų",
        "lv": "Noņemt autortiesības no ekrānuzņēmumiem", "nl": "Copyright van schermafbeeldingen verwijderen",
        "pl": "Usuń prawa autorskie ze zrzutów", "pt": "Remover direitos de autor das capturas",
        "ptbr": "Remover direitos autorais das capturas", "tr": "Ekran görüntülerinden telif hakkını kaldır",
        "zhcn": "移除截图中的版权信息", "zhtw": "移除螢幕截圖中的版權資訊",
    },
    "Блокировать отключение видеозаписи": {
        "ua": "Блокувати вимкнення відеозапису", "be": "Блакаваць адключэнне відэазапісу",
        "de": "Videoaufnahme-Deaktivierung blockieren", "en": "Block disabling of video recording",
        "engb": "Block disabling of video recording", "es": "Bloquear desactivación de grabación de vídeo",
        "es419": "Bloquear desactivación de grabación de vídeo", "et": "Blokeeri videosalvestuse keelamine",
        "fr": "Bloquer la désactivation de l'enregistrement vidéo",
        "frca": "Bloquer la désactivation de l'enregistrement vidéo",
        "id": "Blokir penonaktifan perekaman video", "it": "Blocca disattivazione registrazione video",
        "jp": "動画撮影の無効化をブロック", "kk": "Бейнежазбаны өшіруді блоктау",
        "kr": "동영상 녹화 비활성화 차단", "lt": "Blokuoti vaizdo įrašymo išjungimą",
        "lv": "Bloķēt videoieraksta atspējošanu", "nl": "Uitschakelen video-opname blokkeren",
        "pl": "Blokuj wyłączanie nagrywania wideo", "pt": "Bloquear desativação da gravação de vídeo",
        "ptbr": "Bloquear desativação da gravação de vídeo", "tr": "Video kaydının devre dışı bırakılmasını engelle",
        "zhcn": "阻止禁用视频录制", "zhtw": "封鎖停用影片錄製",
    },
    "Блокировать отключение скриншотов": {
        "ua": "Блокувати вимкнення скриншотів", "be": "Блакаваць адключэнне скрыншотаў",
        "de": "Screenshot-Deaktivierung blockieren", "en": "Block disabling of screenshots",
        "engb": "Block disabling of screenshots", "es": "Bloquear desactivación de capturas",
        "es419": "Bloquear desactivación de capturas", "et": "Blokeeri ekraanipiltide keelamine",
        "fr": "Bloquer la désactivation des captures", "frca": "Bloquer la désactivation des captures",
        "id": "Blokir penonaktifan tangkapan layar", "it": "Blocca disattivazione screenshot",
        "jp": "スクリーンショットの無効化をブロック", "kk": "Скриншоттарды өшіруді блоктау",
        "kr": "스크린샷 비활성화 차단", "lt": "Blokuoti ekrano kopijų išjungimą",
        "lv": "Bloķēt ekrānuzņēmumu atspējošanu", "nl": "Uitschakelen schermafbeeldingen blokkeren",
        "pl": "Blokuj wyłączanie zrzutów ekranu", "pt": "Bloquear desativação das capturas",
        "ptbr": "Bloquear desativação das capturas", "tr": "Ekran görüntüsünün devre dışı bırakılmasını engelle",
        "zhcn": "阻止禁用截图", "zhtw": "封鎖停用螢幕截圖",
    },
    "Не показывать диалог про перезагрузку": {
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
    },
    "Не останавливать игру в фоне": {
        "ua": "Не зупиняти гру у фоні", "be": "Не спыняць гульню у фоне",
        "de": "Spiel im Hintergrund nicht anhalten", "en": "Do not suspend the game in the background",
        "engb": "Do not suspend the game in the background", "es": "No suspender el juego en segundo plano",
        "es419": "No suspender el juego en segundo plano", "et": "Ära peata mängu taustal",
        "fr": "Ne pas suspendre le jeu en arrière-plan", "frca": "Ne pas suspendre le jeu en arrière-plan",
        "id": "Jangan hentikan game di latar belakang", "it": "Non sospendere il gioco in background",
        "jp": "バックグラウンドでゲームを停止しない", "kk": "Ойынды фонда тоқтатпау",
        "kr": "백그라운드에서 게임 중지 안 함", "lt": "Nesustabdyti žaidimo fone",
        "lv": "Neapturēt spēli fonā", "nl": "Spel niet op de achtergrond onderbreken",
        "pl": "Nie zawieszaj gry w tle", "pt": "Não suspender o jogo em segundo plano",
        "ptbr": "Não suspender o jogo em segundo plano", "tr": "Oyunu arka planda askıya alma",
        "zhcn": "不在后台暂停游戏", "zhtw": "不在背景暫停遊戲",
    },
    "Блокировать изменение частот APU": {
        "ua": "Блокувати зміну частот APU", "be": "Блакаваць змену частот APU",
        "de": "Änderung der APU-Frequenzen blockieren", "en": "Block APU frequency changes",
        "engb": "Block APU frequency changes", "es": "Bloquear cambios de frecuencia de APU",
        "es419": "Bloquear cambios de frecuencia de APU", "et": "Blokeeri APU sageduse muutmine",
        "fr": "Bloquer la modification des fréquences APU", "frca": "Bloquer la modification des fréquences APU",
        "id": "Blokir perubahan frekuensi APU", "it": "Blocca modifica frequenze APU",
        "jp": "APU 周波数の変更をブロック", "kk": "APU жиіліктерін өзгертуді блоктау",
        "kr": "APU 주파수 변경 차단", "lt": "Blokuoti APU dažnių keitimą",
        "lv": "Bloķēt APU frekvenču maiņu", "nl": "Wijziging APU-frequenties blokkeren",
        "pl": "Blokuj zmiany częstotliwości APU", "pt": "Bloquear alteração das frequências da APU",
        "ptbr": "Bloquear alteração das frequências da APU", "tr": "APU frekansı değişikliklerini engelle",
        "zhcn": "阻止更改 APU 频率", "zhtw": "封鎖變更 APU 頻率",
    },
}


def forms(key: str) -> list[str]:
    out = [unicodedata.normalize("NFC", key)]
    nfd = unicodedata.normalize("NFD", key)
    if nfd != out[0]:
        out.append(nfd)
    return out


def main() -> None:
    langs = json.loads((ROOT / "data" / "languages.json").read_text(encoding="utf-8"))
    missing = {k: sorted(set(langs) - set(v)) for k, v in T.items()}
    missing = {k: v for k, v in missing.items() if v}
    if missing:
        sys.exit(f"missing translations: {missing}")

    wb = openpyxl.load_workbook(DICT)
    ws = wb["Translations"]
    col = {h: i + 1 for i, h in enumerate(c.value for c in ws[1]) if h}

    rows: dict[str, int] = {}
    for r in range(2, ws.max_row + 1):
        o = ws.cell(r, col["Original"]).value
        if o is not None:
            rows.setdefault(str(o), r)

    added = updated = 0
    for key, per_lang in T.items():
        for form in forms(key):
            if form in rows:
                r = rows[form]
                for lang, text in per_lang.items():
                    if ws.cell(r, col[lang]).value != text:
                        ws.cell(r, col[lang], text)
                        updated += 1
            else:
                r = ws.max_row + 1
                ws.cell(r, col["Original"], form)
                for lang, text in per_lang.items():
                    ws.cell(r, col[lang], text)
                rows[form] = r
                added += 1
    wb.save(DICT)
    print(f"workbook: {added} rows added, {updated} cells updated, {ws.max_row - 1} total")

    # Every language CSV: refresh from the workbook so the pipeline invariant holds.
    for lang in langs:
        subprocess.run([sys.executable, "-c",
                        f"import sys;sys.path.insert(0,'.');"
                        f"from src.main import export_language;export_language('{lang}')"],
                       cwd=ROOT, check=True, capture_output=True)
    print(f"exported {len(langs)} CSVs")


if __name__ == "__main__":
    main()
