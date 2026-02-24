
import pandas as pd
from ics import Calendar, Event
from datetime import datetime
import pytz
import re

def excel_to_ics(input_file, output_file, year, month):
    # Excel dosyasını oku (başlıksız olarak, çünkü grid yapısı var)
    df = pd.read_excel(input_file, header=None)
    
    # Takvim nesnesini oluştur
    cal = Calendar()
    timezone = pytz.timezone("Europe/Istanbul")
    
    # Grid taraması
    # Mantık: Eğer bir hücrede 1-31 arası sayı varsa, bu bir gündür.
    # Bu günün etkinliği genelde bir alt satırdaki aynı sütundadır.
    
    # Ay sonundaki bir sonraki ay günlerini (ör: 1,2,3,4 Nisan) tespit etmek için
    prev_day = 0
    
    for row_idx, row in df.iterrows():
        for col_idx, cell_value in enumerate(row):
            
            # 1. Hücrenin gün sayısı olup olmadığını kontrol et
            try:
                day = int(cell_value)
                if not (1 <= day <= 31):
                    continue
            except (ValueError, TypeError):
                continue
            
            # Ay sonundan sonra gelen küçük sayılar (1,2,3,4...) bir sonraki aya aittir.
            # Örnek: 29,30,31,1,2,3,4 -> buradaki 1-4 aslında Nisan günleri.
            if day < 10 and prev_day > 25:
                continue
            
            prev_day = max(prev_day, day)
            
            # 2. Eğer gün bulduysak, etkinliği hemen altındaki satırda ara
            try:
                event_text = df.iloc[row_idx + 1, col_idx]
            except IndexError:
                continue # Alt satır yoksa geç
            
            # 3. Etkinlik metni boş değilse (NaN veya boş string değilse) işle
            if pd.notna(event_text) and str(event_text).strip():
                event_text = str(event_text).strip()

                # Clean up multiple spaces
                event_text = re.sub(r'\s+', ' ', event_text).strip()

                # Skip if the text is just punctuation (like the comma on Feb 17)
                if len(event_text) < 2:
                    continue
                
                # Etkinlik Tarihini Ayarla (Tüm gün etkinliği)
                event_date = datetime(year, month, day)
                
                e = Event()
                e.name = event_text.split("OC:")[0].strip() # OC kısmını başlıktan ayır
                e.begin = event_date.strftime("%Y-%m-%d")
                e.make_all_day()
                
                # Açıklama kısmına OC bilgisini ekle (boş "OC:" bırakma)
                if "OC:" in event_text:
                    oc_part = event_text.split("OC:", 1)[1].strip()
                    if oc_part:
                        e.description = "OC: " + oc_part
                
                cal.events.add(e)

    # Dosyayı kaydet
    with open(output_file, 'w', encoding='utf-8') as f:
        f.writelines(cal.serialize_iter())
    
    print(f"Başarılı! '{output_file}' dosyası oluşturuldu.")

# --- KULLANIM ---
# Buraya kendi dosya adını ve takvim ayını gir
input_excel = "March Calendar.xlsx" # Excel dosyanın adı
output_ics = "ESN_Eventleri_march.ics"

# Şubat 2026 için çalıştırıyoruz
excel_to_ics(input_excel, output_ics, year=2026, month=3)