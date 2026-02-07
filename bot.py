import telebot
from telebot import types
from datetime import datetime, timedelta
import pandas as pd

# Bot token'ı
BOT_TOKEN = '8524104921:AAHU41XE_cNVDyxJIGRbolyuTLkpYgfN29A'

EXCEL_DOSYASI = 'vardiya.xlsx'

def yukle_vardiya_listesi(excel_dosyasi=EXCEL_DOSYASI):
    try:
        df = pd.read_excel(excel_dosyasi, header=0)
    except FileNotFoundError:
        print(f"HATA: {excel_dosyasi} dosyası bulunamadı!")
        return {}
    except Exception as e:
        print(f"Excel okuma hatası: {e}")
        return {}
    
    tarih_sutunlari = {}
    for col in df.columns:
        if isinstance(col, str) and '.' in col and len(col.split('.')) == 3:
            try:
                gun, ay, yil = map(int, col.split('.'))
                standart_tarih = f"{yil}-{ay:02d}-{gun:02d}"
                tarih_sutunlari[standart_tarih] = col
            except ValueError:
                pass
    
    print(f"{len(tarih_sutunlari)} tarih sütunu bulundu.")
    
    vardiya_dict = {}
    vardiya_kod_map = {
        'SA08': 'sabah',
        'AK08': 'akşam',
        'GE08': 'gece',
        'OFF': 'off'
    }
    
    for _, row in df.iterrows():
        isim = str(row.get('İSİM SOYİSİM', '')).strip()
        if not isim or 'İSİM SOYİSİM' in isim.upper() or isim == '':
            continue
        
        for standart_tarih, orij_sutun in tarih_sutunlari.items():
            kod = str(row.get(orij_sutun, '')).strip().upper()
            if not kod:
                continue
            
            vardiya_tipi = vardiya_kod_map.get(kod)
            if vardiya_tipi is None:
                continue
            
            if standart_tarih not in vardiya_dict:
                vardiya_dict[standart_tarih] = {'sabah': [], 'akşam': [], 'gece': [], 'off': []}
            
            vardiya_dict[standart_tarih][vardiya_tipi].append(isim)
    
    for tarih, vardiyalar in vardiya_dict.items():
        for tip, kisiler in vardiyalar.items():
            if kisiler and tip != 'off' and len(kisiler) != 2:
                print(f"Uyarı: {tarih} {tip} vardiyasında {len(kisiler)} kişi: {kisiler}")
    
    return vardiya_dict

vardiya_listesi = yukle_vardiya_listesi()
print(f"Toplam {len(vardiya_listesi)} tarih yüklendi.")

bot = telebot.TeleBot(BOT_TOKEN)

# Bilgi mesajı (sadece /start ile veya ilk boş mesajda)
@bot.message_handler(commands=['start'])
def start(message):
    bilgi_mesaji = (
        "Vardiyada kim var sorgusu botu.\n\n"
        "Komutlar:\n"
        "/start - Botu çalıştırır, komutlar ve butonlar gelir.\n"
        "/yenile veya /temizle - Sohbeti temizlemek için. \n\n"
        "Tarih seçmek için butonları kullanın. Tarih girmek: 19 şubat ."
    )
    bot.reply_to(message, bilgi_mesaji)
    
    markup = types.InlineKeyboardMarkup(row_width=3)
    markup.add(
        types.InlineKeyboardButton("Bugün", callback_data="bugun"),
        types.InlineKeyboardButton("Yarın", callback_data="yarin"),
        types.InlineKeyboardButton("Tarih Seç", callback_data="tarih_sec")
    )
    bot.send_message(message.chat.id, "Vardiya / izin sorgula:", reply_markup=markup)

# Yenile / temizle komutu
@bot.message_handler(commands=['yenile', 'temizle'])
def yenile(message):
    try:
        # Son 100 mesajı silmeye çalış
        for i in range(1, 101):
            try:
                bot.delete_message(message.chat.id, message.message_id - i)
            except telebot.apihelper.ApiTelegramException as e:
                if "message to delete not found" in str(e):
                    break  # daha eski mesaj yok
                else:
                    print(f"Silme hatası: {e}")
                    break
    except Exception as e:
        print(f"Genel silme hatası: {e}")
    
    # Yeni başlangıç mesajı
    markup = types.InlineKeyboardMarkup(row_width=3)
    markup.add(
        types.InlineKeyboardButton("Bugün", callback_data="bugun"),
        types.InlineKeyboardButton("Yarın", callback_data="yarin"),
        types.InlineKeyboardButton("Tarih Seç", callback_data="tarih_sec")
    )
    bot.send_message(message.chat.id, "Sohbet yenilendi! Eski mesajlar silindi (mümkün olanlar).", reply_markup=markup)

# Tarih parse
def parse_tarih(mesaj):
    bugun = datetime.now().date()
    mesaj = mesaj.lower()
    if 'yarın' in mesaj:
        return (bugun + timedelta(days=1)).strftime('%Y-%m-%d')
    elif 'bugün' in mesaj:
        return bugun.strftime('%Y-%m-%d')
    else:
        try:
            parcalar = mesaj.split()
            gun_str = parcalar[0]
            ay_str = parcalar[1]
            ay_dict = {
                'ocak': 1, 'şubat': 2, 'mart': 3, 'nisan': 4, 'mayıs': 5, 'haziran': 6,
                'temmuz': 7, 'ağustos': 8, 'eylül': 9, 'ekim': 10, 'kasım': 11, 'aralık': 12
            }
            ay = ay_dict.get(ay_str)
            if ay is None:
                return None
            gun = int(gun_str)
            return f"2026-{ay:02d}-{gun:02d}"
        except:
            return None

# Vardiya butonları
def goster_vardiya_butonlari(chat_id, tarih):
    markup = types.InlineKeyboardMarkup(row_width=4)
    markup.add(
        types.InlineKeyboardButton("Sabah", callback_data=f"{tarih}|sabah"),
        types.InlineKeyboardButton("Akşam", callback_data=f"{tarih}|akşam"),
        types.InlineKeyboardButton("Gece", callback_data=f"{tarih}|gece"),
        types.InlineKeyboardButton("Off / İzin", callback_data=f"{tarih}|off")
    )
    bot.send_message(chat_id, f"**{tarih}** için hangi durum?", reply_markup=markup)

@bot.callback_query_handler(func=lambda call: True)
def callback(call):
    if call.data in ["bugun", "yarin"]:
        mod = call.data
        tarih = datetime.now().date().strftime('%Y-%m-%d') if mod == "bugun" else (datetime.now().date() + timedelta(days=1)).strftime('%Y-%m-%d')
        goster_vardiya_butonlari(call.message.chat.id, tarih)
        bot.answer_callback_query(call.id)

    elif call.data == "tarih_sec":
        bot.answer_callback_query(call.id)
        msg = bot.send_message(call.message.chat.id, "Tarihi yaz (ör: 9 şubat)")
        bot.register_next_step_handler(msg, tarih_alindi)

    else:
        try:
            tarih, vardiya = call.data.split("|")
            if tarih in vardiya_listesi and vardiya in vardiya_listesi[tarih]:
                kisiler = ", ".join(vardiya_listesi[tarih][vardiya])
                if vardiya == 'off':
                    baslik = "İzin / Off olanlar"
                else:
                    baslik = f"{vardiya.capitalize()} vardiyasında"
                cevap = f"{tarih} **{baslik}**: {kisiler or 'Kimse yok'}"
                bot.edit_message_text(cevap, call.message.chat.id, call.message.message_id)
            else:
                bot.edit_message_text("Bu tarih veya durum için bilgi yok.", call.message.chat.id, call.message.message_id)
            bot.answer_callback_query(call.id)
        except Exception as e:
            bot.answer_callback_query(call.id, "Hata oluştu")
            print(f"Callback hatası: {e}")

def tarih_alindi(message):
    tarih = parse_tarih(message.text)
    if tarih and tarih in vardiya_listesi:
        goster_vardiya_butonlari(message.chat.id, tarih)
    else:
        bot.reply_to(message, "Geçersiz tarih veya veri yok. Örnek: 9 şubat")

# Bot başlat
print("Bot Telegram'a bağlanıyor ve dinliyor... (Ctrl+C ile durdur)")
bot.polling(none_stop=True)