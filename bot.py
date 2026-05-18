import os
import re
from datetime import datetime, timedelta, date
from zoneinfo import ZoneInfo

import pandas as pd
import telebot
from telebot import types


# =========================
# AYARLAR
# =========================

BOT_TOKEN = os.getenv("BOT_TOKEN")
EXCEL_DOSYASI = os.getenv("EXCEL_DOSYASI", "vardiya.xlsx")
VARSAYILAN_YIL = int(os.getenv("VARSAYILAN_YIL", "2026"))
TIMEZONE = ZoneInfo(os.getenv("TIMEZONE", "Europe/Istanbul"))

if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN bulunamadı. Render Environment Variables içine BOT_TOKEN eklemelisin.")


# =========================
# YARDIMCI FONKSİYONLAR
# =========================

def turkce_temizle(metin):
    """
    Türkçe karakterleri sadeleştirir.
    Örnek: 'şubat' -> 'subat', 'bugün' -> 'bugun'
    """
    metin = str(metin).strip().lower()
    cevir = {
        "ç": "c",
        "ğ": "g",
        "ı": "i",
        "ö": "o",
        "ş": "s",
        "ü": "u",
        "İ": "i",
    }

    for eski, yeni in cevir.items():
        metin = metin.replace(eski, yeni)

    return metin


def bugunun_tarihi():
    return datetime.now(TIMEZONE).date()


def standart_tarih_formatla(tarih_obj):
    return tarih_obj.strftime("%Y-%m-%d")


def ekranda_tarih_formatla(tarih_str):
    try:
        tarih_obj = datetime.strptime(tarih_str, "%Y-%m-%d").date()
        return tarih_obj.strftime("%d.%m.%Y")
    except Exception:
        return tarih_str


def excel_sutunundan_tarih_al(sutun):
    """
    Excel'deki tarih sütunlarını yakalar.
    Şunları destekler:
    - 1.05.2026
    - 01.05.2026
    - 1/05/2026
    - Excel tarih formatı
    """

    if isinstance(sutun, datetime):
        return sutun.date().strftime("%Y-%m-%d")

    if isinstance(sutun, date):
        return sutun.strftime("%Y-%m-%d")

    metin = str(sutun).strip()

    # Excel bazen başlıkları "2026-05-01 00:00:00" gibi okuyabilir
    try:
        tarih = pd.to_datetime(metin, dayfirst=True, errors="coerce")
        if pd.notna(tarih):
            return tarih.date().strftime("%Y-%m-%d")
    except Exception:
        pass

    metin = metin.replace("/", ".").replace("-", ".")
    eslesme = re.match(r"^(\d{1,2})\.(\d{1,2})\.(\d{4})$", metin)

    if not eslesme:
        return None

    gun = int(eslesme.group(1))
    ay = int(eslesme.group(2))
    yil = int(eslesme.group(3))

    try:
        return date(yil, ay, gun).strftime("%Y-%m-%d")
    except ValueError:
        return None


def isim_sutunu_bul(df):
    """
    Excel'deki isim sütununu bulur.
    Normalde sütun adı: İSİM SOYİSİM
    Bulamazsa ilk sütunu isim sütunu kabul eder.
    """
    for col in df.columns:
        temiz = turkce_temizle(col)
        if temiz in ["isim soyisim", "isim soyisim", "ad soyad", "ad soyadi"]:
            return col

    return df.columns[0]


# =========================
# EXCEL OKUMA
# =========================

def yukle_vardiya_listesi(excel_dosyasi=EXCEL_DOSYASI):
    try:
        df = pd.read_excel(excel_dosyasi, header=0)
    except FileNotFoundError:
        print(f"HATA: {excel_dosyasi} dosyası bulunamadı!")
        return {}
    except Exception as e:
        print(f"Excel okuma hatası: {e}")
        return {}

    if df.empty:
        print("Excel boş görünüyor.")
        return {}

    isim_col = isim_sutunu_bul(df)

    tarih_sutunlari = {}

    for col in df.columns:
        standart_tarih = excel_sutunundan_tarih_al(col)
        if standart_tarih:
            tarih_sutunlari[standart_tarih] = col

    print(f"{len(tarih_sutunlari)} tarih sütunu bulundu.")

    vardiya_kod_map = {
        "SA08": "sabah",
        "AK08": "aksam",
        "GE08": "gece",
        "OFF": "off",
    }

    vardiya_dict = {}

    for _, row in df.iterrows():
        isim = str(row.get(isim_col, "")).strip()

        if not isim or isim.lower() == "nan":
            continue

        if "isim" in turkce_temizle(isim):
            continue

        for standart_tarih, orijinal_sutun in tarih_sutunlari.items():
            kod = str(row.get(orijinal_sutun, "")).strip().upper()

            if not kod or kod == "NAN":
                continue

            vardiya_tipi = vardiya_kod_map.get(kod)

            if vardiya_tipi is None:
                continue

            if standart_tarih not in vardiya_dict:
                vardiya_dict[standart_tarih] = {
                    "sabah": [],
                    "aksam": [],
                    "gece": [],
                    "off": [],
                }

            vardiya_dict[standart_tarih][vardiya_tipi].append(isim)

    print(f"Toplam {len(vardiya_dict)} tarih yüklendi.")
    return vardiya_dict


vardiya_listesi = yukle_vardiya_listesi()


# =========================
# TELEGRAM BOT
# =========================

bot = telebot.TeleBot(BOT_TOKEN)


def ana_menu_gonder(chat_id):
    markup = types.InlineKeyboardMarkup(row_width=3)

    markup.add(
        types.InlineKeyboardButton("Bugün", callback_data="bugun"),
        types.InlineKeyboardButton("Yarın", callback_data="yarin"),
        types.InlineKeyboardButton("Tarih Seç", callback_data="tarih_sec"),
    )

    bot.send_message(chat_id, "Vardiya / izin sorgula:", reply_markup=markup)


@bot.message_handler(commands=["start"])
def start(message):
    bilgi_mesaji = (
        "Vardiyada kim var sorgusu botu.\n\n"
        "Komutlar:\n"
        "/start - Botu başlatır.\n"
        "/guncelle - Excel dosyasını yeniden okur.\n"
        "/yenile veya /temizle - Menüyü tekrar getirir.\n\n"
        "Tarih örnekleri:\n"
        "19 mayıs\n"
        "19.05\n"
        "19.05.2026\n"
        "bugün\n"
        "yarın"
    )

    bot.reply_to(message, bilgi_mesaji)
    ana_menu_gonder(message.chat.id)


@bot.message_handler(commands=["yenile", "temizle"])
def yenile(message):
    bot.send_message(message.chat.id, "Menü yenilendi.")
    ana_menu_gonder(message.chat.id)


@bot.message_handler(commands=["guncelle"])
def guncelle(message):
    global vardiya_listesi

    bot.send_message(message.chat.id, "Excel yeniden okunuyor...")

    vardiya_listesi = yukle_vardiya_listesi()

    if vardiya_listesi:
        bot.send_message(message.chat.id, f"Excel güncellendi. Toplam {len(vardiya_listesi)} tarih yüklendi.")
    else:
        bot.send_message(message.chat.id, "Excel okunamadı veya veri bulunamadı.")


def parse_tarih(mesaj):
    """
    Kullanıcının yazdığı tarihi çözer.

    Desteklenenler:
    - bugün
    - yarın
    - 19 mayıs
    - 19 mayis
    - 19.05
    - 19.05.2026
    """

    temiz = turkce_temizle(mesaj)
    bugun = bugunun_tarihi()

    if "yarin" in temiz:
        return standart_tarih_formatla(bugun + timedelta(days=1))

    if "bugun" in temiz:
        return standart_tarih_formatla(bugun)

    # 19.05 veya 19.05.2026
    sayili_tarih = re.search(r"(\d{1,2})[./-](\d{1,2})(?:[./-](\d{2,4}))?", temiz)

    if sayili_tarih:
        gun = int(sayili_tarih.group(1))
        ay = int(sayili_tarih.group(2))
        yil_raw = sayili_tarih.group(3)

        if yil_raw:
            yil = int(yil_raw)
            if yil < 100:
                yil += 2000
        else:
            yil = VARSAYILAN_YIL

        try:
            return date(yil, ay, gun).strftime("%Y-%m-%d")
        except ValueError:
            return None

    ay_dict = {
        "ocak": 1,
        "subat": 2,
        "mart": 3,
        "nisan": 4,
        "mayis": 5,
        "haziran": 6,
        "temmuz": 7,
        "agustos": 8,
        "eylul": 9,
        "ekim": 10,
        "kasim": 11,
        "aralik": 12,
    }

    parcalar = temiz.split()

    gun = None
    ay = None
    yil = VARSAYILAN_YIL

    for parca in parcalar:
        if parca.isdigit():
            sayi = int(parca)
            if 1 <= sayi <= 31 and gun is None:
                gun = sayi
            elif sayi >= 2020:
                yil = sayi

        if parca in ay_dict:
            ay = ay_dict[parca]

    if gun and ay:
        try:
            return date(yil, ay, gun).strftime("%Y-%m-%d")
        except ValueError:
            return None

    return None


def goster_vardiya_butonlari(chat_id, tarih):
    if tarih not in vardiya_listesi:
        bot.send_message(
            chat_id,
            f"{ekranda_tarih_formatla(tarih)} için Excel'de veri bulunamadı."
        )
        return

    markup = types.InlineKeyboardMarkup(row_width=4)

    markup.add(
        types.InlineKeyboardButton("Sabah", callback_data=f"{tarih}|sabah"),
        types.InlineKeyboardButton("Akşam", callback_data=f"{tarih}|aksam"),
        types.InlineKeyboardButton("Gece", callback_data=f"{tarih}|gece"),
        types.InlineKeyboardButton("Off / İzin", callback_data=f"{tarih}|off"),
    )

    bot.send_message(
        chat_id,
        f"{ekranda_tarih_formatla(tarih)} için hangi durum?",
        reply_markup=markup
    )


@bot.callback_query_handler(func=lambda call: True)
def callback(call):
    try:
        if call.data == "bugun":
            tarih = bugunun_tarihi().strftime("%Y-%m-%d")
            goster_vardiya_butonlari(call.message.chat.id, tarih)
            bot.answer_callback_query(call.id)
            return

        if call.data == "yarin":
            tarih = (bugunun_tarihi() + timedelta(days=1)).strftime("%Y-%m-%d")
            goster_vardiya_butonlari(call.message.chat.id, tarih)
            bot.answer_callback_query(call.id)
            return

        if call.data == "tarih_sec":
            bot.answer_callback_query(call.id)
            msg = bot.send_message(call.message.chat.id, "Tarihi yaz. Örnek: 19 mayıs veya 19.05")
            bot.register_next_step_handler(msg, tarih_alindi)
            return

        tarih, vardiya = call.data.split("|")

        vardiya_basliklari = {
            "sabah": "Sabah vardiyasında",
            "aksam": "Akşam vardiyasında",
            "gece": "Gece vardiyasında",
            "off": "Off / izin olanlar",
        }

        kisiler = vardiya_listesi.get(tarih, {}).get(vardiya, [])

        if kisiler:
            kisi_metni = "\n".join([f"- {kisi}" for kisi in kisiler])
        else:
            kisi_metni = "Kimse yok"

        cevap = (
            f"{ekranda_tarih_formatla(tarih)}\n"
            f"{vardiya_basliklari.get(vardiya, vardiya)}:\n"
            f"{kisi_metni}"
        )

        bot.edit_message_text(
            cevap,
            call.message.chat.id,
            call.message.message_id
        )

        bot.answer_callback_query(call.id)

    except Exception as e:
        print(f"Callback hatası: {e}")
        bot.answer_callback_query(call.id, "Hata oluştu.")


def tarih_alindi(message):
    tarih = parse_tarih(message.text)

    if not tarih:
        bot.reply_to(message, "Tarihi anlayamadım. Örnek: 19 mayıs veya 19.05")
        return

    goster_vardiya_butonlari(message.chat.id, tarih)


@bot.message_handler(func=lambda message: True)
def normal_mesaj(message):
    tarih = parse_tarih(message.text)

    if tarih:
        goster_vardiya_butonlari(message.chat.id, tarih)
    else:
        bot.reply_to(message, "Tarih yazabilir veya /start ile menüyü açabilirsin. Örnek: 19 mayıs")


# =========================
# BOTU BAŞLAT
# =========================

print("Bot Telegram'a bağlanıyor ve dinliyor...")
bot.infinity_polling(timeout=60, long_polling_timeout=60)
