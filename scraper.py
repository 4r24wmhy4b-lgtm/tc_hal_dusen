import os
import requests
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright
import pandas as pd
import tempfile
from bs4 import BeautifulSoup

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
    """HTML dosyasını BeautifulSoup ile oku"""
    # Dosyayı BINARY olarak oku
    with open(file_path, 'rb') as f:
        raw_content = f.read()
    
    # Encoding'i otomatik tespit et
    if raw_content.startswith(b'\xff\xfe'):
        html_content = raw_content.decode('utf-16')
    elif raw_content.startswith(b'\xfe\xff'):
        html_content = raw_content.decode('utf-16')
    elif raw_content.startswith(b'\xef\xbb\xbf'):
        html_content = raw_content.decode('utf-8-sig')
    else:
        try:
            html_content = raw_content.decode('utf-8')
        except UnicodeDecodeError:
            try:
                html_content = raw_content.decode('windows-1254')
            except UnicodeDecodeError:
                html_content = raw_content.decode('latin-1')
    
    # BeautifulSoup ile parse et
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Tabloyu bul
    table = soup.find('table')
    
    # Satırları topla
    rows = []
    for tr in table.find_all('tr')[2:]:  # İlk 2 satır başlık, atla
        cells = tr.find_all('td')
        if len(cells) >= 6:
            try:
                urun_adi = cells[0].get_text(strip=True)
                urun_cinsi = cells[1].get_text(strip=True)
                urun_turu = cells[2].get_text(strip=True)
                
                # Fiyatı temizle (span içindeki sayıyı al)
                fiyat_span = cells[3].find('span')
                if fiyat_span:
                    fiyat_text = fiyat_span.get_text(strip=True)
                else:
                    fiyat_text = cells[3].get_text(strip=True)
                
                # Virgülü noktaya çevir
                fiyat_text = fiyat_text.replace(',', '.')
                ortalama_fiyat = float(fiyat_text)
                
                # İşlem hacmi
                hacim_span = cells[4].find('span')
                if hacim_span:
                    hacim_text = hacim_span.get_text(strip=True)
                else:
                    hacim_text = cells[4].get_text(strip=True)
                hacim_text = hacim_text.replace(',', '.')
                islem_hacmi = float(hacim_text)
                
                birim = cells[5].get_text(strip=True)
                
                rows.append({
                    'Urun_Adi': urun_adi,
                    'Urun_Cinsi': urun_cinsi,
                    'Urun_Turu': urun_turu,
                    'Ortalama_Fiyat': ortalama_fiyat,
                    'Islem_Hacmi': islem_hacmi,
                    'Birim': birim
                })
            except Exception as e:
                print(f"Satır okuma hatası: {e}")
                continue
    
    # DataFrame oluştur
    df = pd.DataFrame(rows)
    df = df.dropna(subset=['Urun_Adi'])
    df = df[df['Urun_Adi'] != '']
    
    # Debug: İlk 5 satırı yazdır
    print(f"\nToplam {len(df)} ürün okundu")
    print("Örnek veriler:")
    print(df.head(10).to_string())
    print(f"\nFiyat aralığı: {df['Ortalama_Fiyat'].min():.2f} - {df['Ortalama_Fiyat'].max():.2f}\n")
    
    return df

def compare_prices():
    # Bugünün tarihi
    today = datetime.now()
    
    # Bülten 1 gün gecikmeli yayımlanıyor
    bulletin_today = today - timedelta(days=1)
    bulletin_yesterday = today - timedelta(days=2)
    
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
    print("\nBugünün bülteni indiriliyor...")
    today_file = download_excel_for_date(today_str)
    print("Dünün bülteni indiriliyor...")
    yesterday_file = download_excel_for_date(yesterday_str)
    
    # Excel dosyalarını oku
    df_today = parse_excel(today_file)
    df_yesterday = parse_excel(yesterday_file)
    
    # Benzersiz anahtar oluştur: Ürün Adı + Cinsi + Türü
    df_today['Key'] = df_today['Urun_Adi'] + '|' + df_today['Urun_Cinsi'] + '|' + df_today['Urun_Turu']
    df_yesterday['Key'] = df_yesterday['Urun_Adi'] + '|' + df_yesterday['Urun_Cinsi'] + '|' + df_yesterday['Urun_Turu']
    
    # Doğru merge: Key'e göre
    merged = df_today[['Key', 'Urun_Adi', 'Urun_Cinsi', 'Urun_Turu', 'Ortalama_Fiyat']].merge(
        df_yesterday[['Key', 'Ortalama_Fiyat']], 
        on='Key', 
        suffixes=('_bugun', '_dun')
    )
    
    print(f"\nEşleşen ürün sayısı: {len(merged)}")
    
    # Fiyatı düşenleri filtrele
    merged['fark'] = merged['Ortalama_Fiyat_bugun'] - merged['Ortalama_Fiyat_dun']
    price_dropped = merged[merged['fark'] < 0].copy()
    price_dropped['yuzde_degisim'] = ((price_dropped['Ortalama_Fiyat_bugun'] - price_dropped['Ortalama_Fiyat_dun']) / price_dropped['Ortalama_Fiyat_dun'] * 100).round(1)
    
    # Mesaj oluştur
    if len(price_dropped) == 0:
        message = "📊 **Hal Fiyat Raporu**\n\n✅ Bugün fiyatı düşen ürün bulunamadı."
    else:
        message = f" **Hal Fiyat Raporu**\n\n"
        message += f"📅 {today_str} vs {yesterday_str}\n\n"
        message += f"**Fiyatı Düşen {len(price_dropped)} Ürün:**\n\n"
        
        # En çok düşenleri sırala
        price_dropped = price_dropped.sort_values('yuzde_degisim')
        
        for idx, row in price_dropped.head(20).iterrows():
            urun = f"{row['Urun_Adi']} ({row['Urun_Cinsi']})"
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
