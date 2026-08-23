import os
import requests
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright
import pandas as pd
import tempfile
import re

def download_excel_for_date(date_str):
    """Belirli bir tarih için Excel (HTML) dosyasını indir"""
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        
        # Siteye git
        page.goto("https://www.hal.gov.tr/Sayfalar/FiyatDetaylari.aspx")
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)
        
        # Tarih kutusunu bul ve temizle
        date_input = page.locator("#ctl00_ctl37_g_7e86b8d6_3aea_47cf_b1c1_939799a091e0_dateControl_dateControlDate")
        date_input.click()
        date_input.fill("")
        date_input.fill(date_str)
        page.wait_for_timeout(500)
        
        # Enter'a bas (tarihi onayla)
        date_input.press("Enter")
        page.wait_for_timeout(3000)
        
        # Tüm sayfaları seç (radyo butonu)
        radio_button = page.locator("#ctl00_ctl37_g_7e86b8d6_3aea_47cf_b1c1_939799a091e0_rblExcelOptions_1")
        radio_button.click()
        page.wait_for_timeout(500)
        
        # Excel'e Çıkar butonuna tıkla ve dosyayı indir
        with page.expect_download() as download_info:
            page.click("#ctl00_ctl37_g_7e86b8d6_3aea_47cf_b1c1_939799a091e0_btnExcel")
        
        download = download_info.value
        temp_file = tempfile.NamedTemporaryFile(suffix=".xls", delete=False)
        download.save_as(temp_file.name)
        
        browser.close()
        return temp_file.name

def parse_excel(file_path):
    """Excel (HTML) dosyasını oku ve DataFrame döndür"""
    # HTML tablosunu oku
    dfs = pd.read_html(file_path)
    
    # İkinci tablo veri tablosu (ilk tablo başlık)
    if len(dfs) >= 2:
        df = dfs[1]
    else:
        df = dfs[0]
    
    # Sütun isimlerini düzelt
    df.columns = ['Urun_Adi', 'Urun_Cinsi', 'Urun_Turu', 'Ortalama_Fiyat', 'Islem_Hacmi', 'Birim']
    
    # Fiyatları sayısal yap (virgülü noktaya çevir)
    df['Ortalama_Fiyat'] = df['Ortalama_Fiyat'].astype(str).str.replace(',', '.', regex=False)
    df['Ortalama_Fiyat'] = pd.to_numeric(df['Ortalama_Fiyat'], errors='coerce')
    
    # İşlem hacmini sayısal yap
    df['Islem_Hacmi'] = df['Islem_Hacmi'].astype(str).str.replace(',', '.', regex=False)
    df['Islem_Hacmi'] = pd.to_numeric(df['Islem_Hacmi'], errors='coerce')
    
    # Boş satırları temizle
    df = df.dropna(subset=['Urun_Adi'])
    df = df[df['Urun_Adi'] != '']
    
    return df

def compare_prices():
    # Bugünün tarihi
    today = datetime.now()
    
    # Bülten 1 gün gecikmeli yayımlanıyor
    # 23.08.2026'da yayımlanan bülten 22.08.2026 verilerini kullanıyor
    # O yüzden "bugün" için dünün tarihini, "dün" için 2 gün önceki tarihi kullanacağız
    
    bulletin_today = today - timedelta(days=1)  # Bugünün bülteni = dünün verisi
    bulletin_yesterday = today - timedelta(days=2)  # Dünün bülteni = 2 gün önceki verisi
    
    # Hafta sonu kontrolü
    if bulletin_today.weekday() == 5:  # Cumartesi
        bulletin_today = bulletin_today - timedelta(days=1)
        bulletin_yesterday = bulletin_today - timedelta(days=1)
    elif bulletin_today.weekday() == 6:  # Pazar
        bulletin_today = bulletin_today - timedelta(days=2)
        bulletin_yesterday = bulletin_today - timedelta(days=1)
    
    today_str = bulletin_today.strftime("%d.%m.%Y")
    yesterday_str = bulletin_yesterday.strftime("%d.%m.%Y")
    
    print(f"Bülten 1 (Bugün): {today_str}")
    print(f"Bülten 2 (Dün): {yesterday_str}")
    
    # Excel dosyalarını indir
    print("Bugünün bülteni indiriliyor...")
    today_file = download_excel_for_date(today_str)
    print("Dünün bülteni indiriliyor...")
    yesterday_file = download_excel_for_date(yesterday_str)
    
    # Excel dosyalarını oku
    df_today = parse_excel(today_file)
    df_yesterday = parse_excel(yesterday_file)
    
    print(f"Bugün: {len(df_today)} ürün, Dün: {len(df_yesterday)} ürün")
    
    # Verileri birleştir
    merged = df_today[['Urun_Adi', 'Ortalama_Fiyat']].merge(
        df_yesterday[['Urun_Adi', 'Ortalama_Fiyat']], 
        on='Urun_Adi', 
        suffixes=('_bugun', '_dun')
    )
    
    # Fiyatı düşenleri filtrele
    merged['fark'] = merged['Ortalama_Fiyat_bugun'] - merged['Ortalama_Fiyat_dun']
    price_dropped = merged[merged['fark'] < 0].copy()
    price_dropped['yuzde_degisim'] = ((price_dropped['Ortalama_Fiyat_bugun'] - price_dropped['Ortalama_Fiyat_dun']) / price_dropped['Ortalama_Fiyat_dun'] * 100).round(1)
    
    # Mesaj oluştur
    if len(price_dropped) == 0:
        message = "📊 **Hal Fiyat Raporu**\n\n✅ Bugün fiyatı düşen ürün bulunamadı."
    else:
        message = f"📉 **Hal Fiyat Raporu**\n\n"
        message += f"📅 {today_str} vs {yesterday_str}\n\n"
        message += f"**Fiyatı Düşen {len(price_dropped)} Ürün:**\n\n"
        
        # En çok düşenleri sırala
        price_dropped = price_dropped.sort_values('yuzde_degisim')
        
        for idx, row in price_dropped.head(20).iterrows():
            urun = row['Urun_Adi']
            fiyat_dun = row['Ortalama_Fiyat_dun']
            fiyat_bugun = row['Ortalama_Fiyat_bugun']
            degisim = row['yuzde_degisim']
            
            message += f"• {urun}: {fiyat_dun:.2f}₺ ➔ **{fiyat_bugun:.2f}₺** ({degisim}%)\n"
        
        if len(price_dropped) > 20:
            message += f"\n_... ve {len(price_dropped) - 20} ürün daha_"
    
    # Dosyaları temizle
    os.unlink(today_file)
    os.unlink(yesterday_file)
    
    return message

def send_telegram(message):
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID")
    
    if not token or not chat_id:
        print("Hata: Telegram ayarları eksik!")
        return

    url = f"https://api.telegram.org/bot{token}/sendMessage"
    data = {
        "chat_id": chat_id,
        "text": message,
        "parse_mode": "Markdown"
    }
    response = requests.post(url, json=data)
    if response.status_code == 200:
        print("Bildirim başarıyla gönderildi!")
    else:
        print(f"Telegram hatası: {response.text}")

if __name__ == "__main__":
    report = compare_prices()
    send_telegram(report)
