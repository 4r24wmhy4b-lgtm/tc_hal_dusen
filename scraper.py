import os
import requests
import time
from datetime import datetime, timedelta, timezone
from playwright.sync_api import sync_playwright
import pandas as pd
import tempfile
from bs4 import BeautifulSoup


def download_excel_for_date(date_str):
    """Belirli bir tarih için Excel (HTML) dosyasını indir"""
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()

        page.goto("https://www.hal.gov.tr/Sayfalar/FiyatDetaylari.aspx")
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)

        date_input = page.locator(
            "#ctl00_ctl37_g_7e86b8d6_3aea_47cf_b1c1_939799a091e0_dateControl_dateControlDate"
        )
        date_input.click()
        date_input.fill("")
        date_input.fill(date_str)
        page.wait_for_timeout(500)
        date_input.press("Enter")
        page.wait_for_timeout(3000)

        radio_button = page.locator(
            "#ctl00_ctl37_g_7e86b8d6_3aea_47cf_b1c1_939799a091e0_rblExcelOptions_1"
        )
        radio_button.click()
        page.wait_for_timeout(500)

        with page.expect_download() as download_info:
            page.click(
                "#ctl00_ctl37_g_7e86b8d6_3aea_47cf_b1c1_939799a091e0_btnExcel"
            )

        download = download_info.value
        temp_file = tempfile.NamedTemporaryFile(suffix=".xls", delete=False)
        download.save_as(temp_file.name)

        browser.close()
        return temp_file.name


def parse_excel(file_path):
    """HTML dosyasını BeautifulSoup ile oku"""
    with open(file_path, 'rb') as f:
        raw_content = f.read()

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

    soup = BeautifulSoup(html_content, 'html.parser')
    table = soup.find('table')

    rows = []
    for tr in table.find_all('tr')[2:]:
        cells = tr.find_all('td')
        if len(cells) >= 6:
            try:
                urun_adi = cells[0].get_text(strip=True)
                urun_cinsi = cells[1].get_text(strip=True)
                urun_turu = cells[2].get_text(strip=True)

                fiyat_span = cells[3].find('span')
                if fiyat_span:
                    fiyat_text = fiyat_span.get_text(strip=True)
                else:
                    fiyat_text = cells[3].get_text(strip=True)

                fiyat_text = fiyat_text.replace(',', '.')
                ortalama_fiyat = float(fiyat_text)

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

    df = pd.DataFrame(rows)
    df = df.dropna(subset=['Urun_Adi'])
    df = df[df['Urun_Adi'] != '']

    print(f"\nToplam {len(df)} ürün okundu")
    print(f"Fiyat aralığı: {df['Ortalama_Fiyat'].min():.2f} - {df['Ortalama_Fiyat'].max():.2f}\n")

    return df


def filter_by_volume(df):
    """İşlem hacmine göre filtrele"""
    filtered_rows = []
    for idx, row in df.iterrows():
        birim = str(row['Birim']).lower()
        hacim = row['Islem_Hacmi']
        
        if 'kg' in birim:
            # Kg birimi: 10.000 kg altı kaldır
            if hacim >= 10000:
                filtered_rows.append(row)
        elif 'adet' in birim or 'bağ' in birim or 'bag' in birim:
            # Adet/Bağ birimi: 100.000 altı kaldır
            if hacim >= 100000:
                filtered_rows.append(row)
        else:
            # Diğer birimler için filtreleme yok
            filtered_rows.append(row)
    
    df_filtered = pd.DataFrame(filtered_rows)
    print(f"Filtreleme sonrası: {len(df_filtered)} ürün (kaldırılan: {len(df) - len(df_filtered)})")
    return df_filtered


def send_telegram(message):
    """Telegram'a mesaj gönder"""
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID")

    if not token or not chat_id:
        print("Hata: Telegram ayarları eksik!")
        return False

    url = f"https://api.telegram.org/bot{token}/sendMessage"
    data = {
        "chat_id": chat_id,
        "text": message,
        "parse_mode": "Markdown"
    }
    response = requests.post(url, json=data)
    if response.status_code == 200:
        print("Bildirim başarıyla gönderildi!")
        return True
    else:
        print(f"Telegram hatası: {response.text}")
        return False


def compare_prices():
    """Fiyatları karşılaştır ve rapor oluştur"""
    # Türkiye saatini kullan (UTC+3)
    tz_tr = timezone(timedelta(hours=3))
    today = datetime.now(tz_tr)
    
    print(f"Şu anki tarih (Türkiye): {today.strftime('%d.%m.%Y %H:%M')}")

    # Site yapısı: "Bülten Tarihi: X (Y Tarihli Veriler)" - X = Y+1
    # 23.08 bülteni → 22.08 verileri
    # 22.08 bülteni → 21.08 verileri
    
    # Bugünün bültenini al (0 gün geri)
    bulletin_today = today
    
    # Dünün bültenini al (1 gün geri)
    bulletin_yesterday = today - timedelta(days=1)

    today_str = bulletin_today.strftime("%d.%m.%Y")
    yesterday_str = bulletin_yesterday.strftime("%d.%m.%Y")

    # Mesajda gösterilecek gerçek veri tarihleri
    data_today = bulletin_today - timedelta(days=1)
    data_yesterday = bulletin_yesterday - timedelta(days=1)
    data_today_str = data_today.strftime("%d.%m.%Y")
    data_yesterday_str = data_yesterday.strftime("%d.%m.%Y")

    print(f"Bülten 1: {today_str} (içerdiği veri: {data_today_str})")
    print(f"Bülten 2: {yesterday_str} (içerdiği veri: {data_yesterday_str})")

    print("\nBugünün bülteni indiriliyor...")
    today_file = download_excel_for_date(today_str)
    print("Dünün bülteni indiriliyor...")
    yesterday_file = download_excel_for_date(yesterday_str)

    df_today = parse_excel(today_file)
    df_yesterday = parse_excel(yesterday_file)

    # İşlem hacmine göre filtrele
    print("\nFiltreleme uygulanıyor...")
    df_today = filter_by_volume(df_today)
    df_yesterday = filter_by_volume(df_yesterday)

    df_today['Key'] = (
        df_today['Urun_Adi'] + '|' +
        df_today['Urun_Cinsi'] + '|' +
        df_today['Urun_Turu']
    )
    df_yesterday['Key'] = (
        df_yesterday['Urun_Adi'] + '|' +
        df_yesterday['Urun_Cinsi'] + '|' +
        df_yesterday['Urun_Turu']
    )

    merged = df_today[['Key', 'Urun_Adi', 'Urun_Cinsi', 'Urun_Turu', 'Ortalama_Fiyat', 'Islem_Hacmi']].merge(
        df_yesterday[['Key', 'Ortalama_Fiyat', 'Islem_Hacmi']],
        on='Key',
        suffixes=('_bugun', '_dun')
    )

    print(f"\nEşleşen ürün sayısı: {len(merged)}")

    merged['fark'] = merged['Ortalama_Fiyat_bugun'] - merged['Ortalama_Fiyat_dun']
    price_dropped = merged[merged['fark'] < 0].copy()
    price_dropped['yuzde_degisim'] = (
        (price_dropped['Ortalama_Fiyat_bugun'] - price_dropped['Ortalama_Fiyat_dun'])
        / price_dropped['Ortalama_Fiyat_dun'] * 100
    ).round(1)

    print(f"Fiyatı düşen ürün sayısı: {len(price_dropped)}")

    # Mesajları oluştur
    if len(price_dropped) == 0:
        message = " **Hal Fiyat Raporu**\n\n✅ Bugün fiyatı düşen ürün bulunamadı."
        send_telegram(message)
    else:
        # Alfabetik sırala
        price_dropped = price_dropped.sort_values('Urun_Adi')

        header = f"📉 **Hal Fiyat Raporu**\n\n"
        header += f"📅 {data_today_str} vs {data_yesterday_str}\n\n"
        header += f"**Fiyatı Düşen {len(price_dropped)} Ürün:**\n\n"

        messages = []
        current_message = header
        current_length = len(header)

        for idx, row in price_dropped.iterrows():
            urun = f"{row['Urun_Adi']} ({row['Urun_Cinsi']})"
            fiyat_dun = row['Ortalama_Fiyat_dun']
            fiyat_bugun = row['Ortalama_Fiyat_bugun']
            degisim = row['yuzde_degisim']

            line = f"• {urun}: {fiyat_dun:.2f}₺ ➔ **{fiyat_bugun:.2f}₺** ({degisim}%)\n"
            line_length = len(line)

            if current_length + line_length > 4000:
                messages.append(current_message)
                current_message = header + line
                current_length = len(header) + line_length
            else:
                current_message += line
                current_length += line_length

        if current_message and current_message != header:
            messages.append(current_message)

        # Tüm mesajları gönder
        for i, msg in enumerate(messages, 1):
            if len(messages) > 1:
                msg = f"**({i}/{len(messages)})**\n\n{msg}"
            send_telegram(msg)
            time.sleep(1)

    os.unlink(today_file)
    os.unlink(yesterday_file)


if __name__ == "__main__":
    compare_prices()
