"""Apply zh-CN dictionary fixes: new SaltyNX dialog keys + fill one blank line.

Row order in every translations CSV matches the workbook order, so new keys are
appended to all 23 CSVs, data/ua.csv and dictionary.xlsx in one pass.
"""
import csv
import json
import shutil
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TR = ROOT / "translations"
LANGS = list(json.loads((ROOT / "data" / "languages.json").read_text(encoding="utf-8")).keys())

# source key -> {lang: translation}
NEW = {
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
    "Форсировать режим полета": {
        "ua": "Форсувати режим польоту", "be": "Фарсіраваць рэжым палёту",
        "de": "Flugmodus erzwingen", "en": "Force flight mode",
        "engb": "Force flight mode", "es": "Forzar modo avión",
        "es419": "Forzar modo avión", "et": "Sunni lennurežiimi",
        "fr": "Forcer le mode avion", "frca": "Forcer le mode avion",
        "id": "Paksa mode pesawat", "it": "Forza modalità aereo",
        "jp": "機内モードを強制", "kk": "Ұшу режимін мәжбүрлеу",
        "kr": "비행기 모드 강제", "lt": "Priverstinis skrydžio režimas",
        "lv": "Piespiedu lidojuma režīms", "nl": "Vliegtuigmodus forceren",
        "pl": "Wymuś tryb samolotowy", "pt": "Forçar modo de voo",
        "ptbr": "Forçar modo avião", "tr": "Uçuş modunu zorla",
        "zhcn": "强制飞行模式", "zhtw": "強制飛航模式",
    },
    "Игнорировать ответ от сервера": {
        "ua": "Ігнорувати відповідь від сервера", "be": "Ігнараваць адказ ад сервера",
        "de": "Serverantwort ignorieren", "en": "Ignore server response",
        "engb": "Ignore server response", "es": "Ignorar respuesta del servidor",
        "es419": "Ignorar respuesta del servidor", "et": "Eira serveri vastust",
        "fr": "Ignorer la réponse du serveur", "frca": "Ignorer la réponse du serveur",
        "id": "Abaikan respons server", "it": "Ignora risposta del server",
        "jp": "サーバー応答を無視", "kk": "Сервер жауабын елемеу",
        "kr": "서버 응답 무시", "lt": "Nepaisyti serverio atsakymo",
        "lv": "Ignorēt servera atbildi", "nl": "Serverantwoord negeren",
        "pl": "Ignoruj odpowiedź serwera", "pt": "Ignorar resposta do servidor",
        "ptbr": "Ignorar resposta do servidor", "tr": "Sunucu yanıtını yoksay",
        "zhcn": "忽略服务器响应", "zhtw": "忽略伺服器回應",
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
    "Не показывать диалоги с ошибками": {
        "ua": "Не показувати діалоги з помилками", "be": "Не паказваць дыялогі з памылкамі",
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
        "zhcn": "不在后台暂停游戏", "zhtw": "不在背景暂停遊戲",
    },
    "Блокировать изменение частот CPU": {
        "ua": "Блокувати зміну частот CPU", "be": "Блакаваць змену частот CPU",
        "de": "Änderung der CPU-Frequenzen blockieren", "en": "Block CPU frequency changes",
        "engb": "Block CPU frequency changes", "es": "Bloquear cambios de frecuencia de CPU",
        "es419": "Bloquear cambios de frecuencia de CPU", "et": "Blokeeri protsessori sageduse muutmine",
        "fr": "Bloquer la modification des fréquences CPU", "frca": "Bloquer la modification des fréquences CPU",
        "id": "Blokir perubahan frekuensi CPU", "it": "Blocca modifica frequenze CPU",
        "jp": "CPU 周波数の変更をブロック", "kk": "CPU жиіліктерін өзгертуді блоктау",
        "kr": "CPU 주파수 변경 차단", "lt": "Blokuoti CPU dažnių keitimą",
        "lv": "Bloķēt CPU frekvenču maiņu", "nl": "Wijziging CPU-frequenties blokkeren",
        "pl": "Blokuj zmiany częstotliwości CPU", "pt": "Bloquear alteração das frequências da CPU",
        "ptbr": "Bloquear alteração das frequências da CPU", "tr": "CPU frekansı değişikliklerini engelle",
        "zhcn": "阻止更改 CPU 频率", "zhtw": "封鎖變更 CPU 頻率",
    },
}

# One existing key is blank in zhcn/zhtw but has real text in en: the "Bunker State
# Press «Looking-Glass»" line of the Shadok parody block (outside shadok.json mapping).
FILL = {
    "  Жак Руксель. Великолепие навыворот.": {"zhcn": "  地堡，镜中世界国营出版社."},
}


def main() -> None:
    missing = [lang for k in NEW for lang in (set(LANGS) - set(NEW[k]))]
    if missing:
        raise SystemExit(f"missing translations for: {sorted(set(missing))}")

    for lang in LANGS:
        path = TR / f"{lang}.csv"
        rows = list(csv.reader(path.open(encoding="utf-8", newline="")))
        header, data = rows[0], rows[1:]
        existing = {r[0] for r in data}
        added = 0
        for key, tr in NEW.items():
            assert key not in existing, key
            data.append([key, tr[lang]])
            added += 1
        for r in data:
            if r[0] in FILL and lang in FILL[r[0]]:
                assert not r[1].strip(), r
                r[1] = FILL[r[0]][lang]
        assert header == ["original", "translation"], (lang, header)
        with path.open("w", encoding="utf-8", newline="") as f:
            csv.writer(f, lineterminator="\n").writerows([header, *data])
        print(f"{lang}: +{added} rows -> {len(data)}")

    # data/ua.csv is the sync source of truth (original, ukrainian)
    ua_path = ROOT / "data" / "ua.csv"
    ua_rows = list(csv.reader(ua_path.open(encoding="utf-8-sig", newline="")))
    ua_existing = {r[0] for r in ua_rows[1:]}
    for key in NEW:
        assert key not in ua_existing, key
        ua_rows.append([key, NEW[key]["ua"]])
    with ua_path.open("w", encoding="utf-8", newline="") as f:
        csv.writer(f, lineterminator="\n").writerows(ua_rows)
    print(f"data/ua.csv: -> {len(ua_rows) - 1} rows")

    # dictionary.xlsx: append rows and set the new translations
    wb_path = ROOT / "data" / "dictionary.xlsx"
    wb = openpyxl.load_workbook(wb_path)
    ws = wb["Translations"]
    headers = [c.value for c in ws[1]]
    col = {h: i + 1 for i, h in enumerate(headers) if h}
    for key, tr in NEW.items():
        r = ws.max_row + 1
        ws.cell(r, col["Original"], key)
        for lang, text in tr.items():
            ws.cell(r, col[lang], text)
    print(f"dictionary.xlsx: -> {ws.max_row - 1} rows")
    wb.save(wb_path)


if __name__ == "__main__":
    main()
