#Doğrulama Kodu
import requests
from bs4 import BeautifulSoup

url = "https://docs.google.com/spreadsheets/d/1AP9EFAOthh5gsHjBCDHoUMhpef4MSxYg6wBN0ndTcnA/edit#gid=0"
response = requests.get(url)
html_content = response.content
soup = BeautifulSoup(html_content, "html.parser")
first_cell = soup.find("td", {"class": "s2"}).text.strip()
if first_cell != "Aktif":
    exit()
first_cell = soup.find("td", {"class": "s1"}).text.strip()
print(first_cell)

import pandas as pd
from io import BytesIO
import re
from colorama import init, Fore, Style
from datetime import datetime
import numpy as np
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font
from openpyxl.styles import Font, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import os
import asyncio
import xlsxwriter
import openpyxl
from openpyxl.styles import Font
import time
from datetime import datetime, timedelta
import shutil
from tqdm import tqdm
import warnings
from selenium.webdriver.chrome.service import Service
from colorama import init, Fore, Style
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from io import BytesIO
import numpy as np
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time
from tqdm import tqdm
import warnings
from colorama import init, Fore, Style
import threading
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import chromedriver_autoinstaller
from concurrent.futures import ThreadPoolExecutor
import subprocess
from selenium.common.exceptions import TimeoutException, WebDriverException
from bs4 import BeautifulSoup
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from datetime import datetime
from datetime import datetime, timedelta
import shutil
import os
from openpyxl import load_workbook
from openpyxl import Workbook
from pathlib import Path
import re
import http.client
import json
import gc
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl.comments import Comment
from selenium.webdriver.chrome.options import Options
from copy import copy
from openpyxl.styles import PatternFill
warnings.filterwarnings("ignore")
pd.options.mode.chained_assignment = None
init(autoreset=True)

print(" ")
print(Fore.GREEN + "Oturum Açma Başarılı Oldu")
print(" /﹋\ ")
print("(҂`_´)")
print(Fore.RED + "<,︻╦╤─ ҉ - -")
print(" /﹋\ ")
print("Mustafa ARI")

#region // Entegrasyondan Önce mi Sonra mı Kontrolü

# Gizli modda Chrome ayarları
chrome_options = Options()
chrome_options.add_argument("--headless")  # Tarayıcıyı ekranda göstermemek için
chrome_options.add_argument("--incognito")  # Gizli modda çalıştırmak için
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

# ChromeDriver hizmeti
service = Service()

# Tarayıcı başlat
driver = webdriver.Chrome(service=service, options=chrome_options)

# Kullanıcı bilgileri
username = "mustafa_kod@haydigiy.com"
password = "123456"

# URL'ler
login_url = "https://www.siparis.haydigiy.com/kullanici-giris/?ReturnUrl=%2Fadmin"
product_list_url = "https://www.siparis.haydigiy.com/admin/product/list/"

try:
    # Giriş sayfasına git
    driver.get(login_url)
    time.sleep(2)  # Sayfanın yüklenmesini bekleyin

    # Giriş bilgilerini doldur
    driver.find_element(By.NAME, "EmailOrPhone").send_keys(username)
    driver.find_element(By.NAME, "Password").send_keys(password)
    driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
    time.sleep(3)  # Giriş sonrası bekleme süresi

    # Ürün listesi sayfasına git
    driver.get(product_list_url)
    time.sleep(5)  # Sayfanın tamamen yüklenmesini bekleyin

    # Bugünün tarihini alın
    current_date = datetime.now().strftime("%d.%m.%Y")
    current_date_single_digit = f"{datetime.now().day}.{datetime.now().strftime('%m.%Y')}"  # Tek basamaklı gün formatı

    # Dinamik yüklenen ürün verilerini çek
    rows = driver.find_elements(By.CSS_SELECTOR, "tr[data-uid]")

    if not rows:
        print(Fore.RED + "Ürün listesi bulunamadı veya doğru yüklenmedi." + Style.RESET_ALL)
    else:
        # Kontrol için bayrak (flag)
        contains_today = False

        for row in rows:
            # Tarih hücresini bulun (5. sütun)
            date_cell = row.find_elements(By.TAG_NAME, "td")[4].text.strip()

            # Çift basamaklı ve tek basamaklı gün kontrolü
            if current_date in date_cell or current_date_single_digit in date_cell:
                contains_today = True
                break  # Bugünün tarihi bulunduysa döngüyü durdur

        # Bayrağa göre mesaj yazdır
        if contains_today:
            print(Fore.GREEN + "Entegrasyondan Sonraki Listeyi Çekiyorsunuz" + Style.RESET_ALL)
        else:
            print(Fore.RED + "Dikkat Entegrasyondan Önceki Listeyi Çekiyorsunuz !" + Style.RESET_ALL)

except Exception as e:
    print(Fore.RED + f"Hata oluştu: {e}" + Style.RESET_ALL)
finally:
    # Tarayıcıyı kapat
    driver.quit()


#endregion

#region // Seçim Yapma Alanı

etiket_secimi = input("Sadece Etiketli Ürünleri mi Çekmek İstiyorsunuz (E/H): ").strip().upper()

# Kullanıcıdan seçim yapılması
secim = input(Fore.YELLOW + "\n1. Firma Kodu Bazlı\n2. Ürün Adında Geçen Bir Kelime ya da Kısım\n3. Kumaş Bazlı\n4. Kalıp Bazlı\n5. Kategori Bazlı" + Fore.LIGHTCYAN_EX + "\n6. 1-3 Arası Aktif Ürünler\n7. Raf Ömrü Girme" + Fore.WHITE + "\nSeçim: ")

if secim == "1":
    kolon_adi = "UrunAdi"
    kullanici_input = input("Firma Kodu (Ör: .1234.): ")
elif secim == "2":
    kolon_adi = "UrunAdi"
    kullanici_input = input("Ürün Adında Geçen Bir Kısım  (Ör: Kareli): ")
elif secim == "3":
    kolon_adi = "Aciklama"
    kullanici_input = input("Kumaş (Ör: Kaşkorse): ")
elif secim == "4":
    kolon_adi = "Aciklama"
    kullanici_input = input("Kalıp (Ör: Dar): ")
elif secim == "5":
    kolon_adi = "Kategori"
    kullanici_input = input("Kategori (Ör: YENİ SEZON): ")

#endregion

#region // Seçim 6 (1-3 Arası Aktif Ürünler)

elif secim == "6":


    # Excel dosyasını indir
    url = "https://www.siparis.haydigiy.com/FaprikaXls/ODJC6P/1/"
    response = requests.get(url)

    if response.status_code == 200:
        with open("1-3 Arası Aktif Ürünler.xlsx", "wb") as dosya:
            dosya.write(response.content)

        
        # Excel dosyasını oku
        df = pd.read_excel("1-3 Arası Aktif Ürünler.xlsx")

        # "UrunAdi" sütunundaki aynı değerlerin "StokAdedi" sütunundaki değerlerini toplayarak yeni bir "ToplamStok" sütunu oluştur
        df["ToplamStok"] = df.groupby("UrunAdi")["StokAdedi"].transform("sum")

        # "UrunAdi", "AlisFiyati", "SatisFiyati" ve "ToplamStok" sütunlarını sakla, "StokAdedi" sütununu sil
        columns_to_keep = ["UrunAdi", "AlisFiyati", "SatisFiyati", "ToplamStok"]
        df = df[columns_to_keep]

        # Sonucu aynı dosya üzerine kaydet
        df.to_excel("1-3 Arası Aktif Ürünler.xlsx", index=False)

        # Aynı olan satırları teke düşür
        df.drop_duplicates(inplace=True)

        # Excel dosyasını açın
        workbook = openpyxl.load_workbook("1-3 Arası Aktif Ürünler.xlsx")
        
        # İlk çalışma sayfasını seçin (Varsayılan olarak "Sheet1")
        sheet = workbook.active

        # Tüm sütunları gezip en uygun genişliği ayarlayın
        for column in sheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)

            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass

            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column_letter].width = adjusted_width

        
        # Değişiklikleri kaydedin
        workbook.save("1-3 Arası Aktif Ürünler.xlsx")

        # Verileri pandas ile tekrar okuyun
        df = pd.read_excel("1-3 Arası Aktif Ürünler.xlsx")

        # Aynı olan satırları teke düşür
        df.drop_duplicates(inplace=True)

        # Değişiklikleri kaydedin
        df.to_excel("1-3 Arası Aktif Ürünler.xlsx", index=False)


        # Excel dosyasını açın
        workbook = openpyxl.load_workbook("1-3 Arası Aktif Ürünler.xlsx")
        
        # İlk çalışma sayfasını seçin (Varsayılan olarak "Sheet1")
        sheet = workbook.active

        # Tüm sütunları gezip en uygun genişliği ayarlayın
        for column in sheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)

            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass

            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column_letter].width = adjusted_width


        # Değişiklikleri kaydedin
        workbook.save("1-3 Arası Aktif Ürünler.xlsx")


    exit()  


#endregion

#region // Seçim 7 (Raf Ömrü Girme)

elif secim == "7":
    try:
        num_parts = int(input("Listeler Kaç Parçaya Bölünsün: "))

        def get_excel_data(url):
            response = requests.get(url)

            if response.status_code == 200:
                # Excel dosyasını oku
                df = pd.read_excel(BytesIO(response.content))
                return df
            else:
                
                return None

        # İlk linkten veriyi al
        url1 = "https://www.siparis.haydigiy.com/FaprikaXls/Q07PJA/1/"
        data1 = get_excel_data(url1)

        # İkinci linkten veriyi al
        url2 = "https://www.siparis.haydigiy.com/FaprikaXls/Q07PJA/2/"
        data2 = get_excel_data(url2)

        # İki veriyi birleştir
        if data1 is not None and data2 is not None:
            merged_data = pd.concat([data1, data2], ignore_index=True)

            # Gereksiz sütunları sil
            columns_to_keep = ["ModelKodu", "UrunAdi", "Resim", "VaryasyonHepsiBuradaKodu"]
            merged_data = merged_data[columns_to_keep]

            # Birleştirilmiş veriyi Excel dosyasına kaydet
            merged_data.to_excel("birlesmis_veri.xlsx", index=False)

            # Birleştirilmiş veriyi oku
            final_data = pd.read_excel("birlesmis_veri.xlsx")

        else:
            pass





        # XML'den Ürün Bilgilerini Çekme ve Temizleme
        xml_url = "https://www.siparis.haydigiy.com/FaprikaXml/SDDI3V/1/"
        response = requests.get(xml_url)
        xml_data = response.text
        soup = BeautifulSoup(xml_data, 'xml')

        product_data = []
        for item in soup.find_all('item'):
            title = item.find('title').text.replace(' - Haydigiy', '')
            product_id = item.find('g:product_type').text if item.find('g:product_type') else None
            product_data.append({'UrunAdi': title, 'ID': product_id})

        df_xml = pd.DataFrame(product_data)

        # Excel ile Birleştirme
        df_calisma_alani = pd.read_excel('birlesmis_veri.xlsx')
        df_merged = pd.merge(df_calisma_alani, df_xml, how='left', left_on='UrunAdi', right_on='UrunAdi')

        # Sonuçları Mevcut Excel Dosyasının Üzerine Kaydetme
        df_merged.to_excel('birlesmis_veri.xlsx', index=False)




        # Excel dosyasını oku
        merged_data = pd.read_excel("birlesmis_veri.xlsx")

        # Boş olmayan "VaryasyonHepsiBuradaKodu" sütunlarına sahip satırları filtrele
        merged_data = merged_data[merged_data["VaryasyonHepsiBuradaKodu"].isna()]

        # Güncellenmiş birleştirilmiş veriyi Excel dosyasına kaydet
        merged_data.to_excel("birlesmis_veri.xlsx", index=False)





        # Excel dosyasını oku
        merged_data = pd.read_excel("birlesmis_veri.xlsx")

        # Resim sütunundaki ".jpeg" ve sonrasını temizleme
        merged_data["Resim"] = merged_data["Resim"].str.replace("\.jpeg.*$", "", regex=True)

        # Resim sütunundaki verilere ".jpeg" eklenmesi
        merged_data["Resim"] = merged_data["Resim"] + ".jpeg"

        # Güncellenmiş birleştirilmiş veriyi Excel dosyasına kaydet
        merged_data.to_excel("birlesmis_veri.xlsx", index=False)






        # Excel dosyasını oku
        merged_data = pd.read_excel("birlesmis_veri.xlsx")

        # Birleşmiş verilerin kopyasını oluştur
        merged_data_copy = merged_data.copy()

        # İstenmeyen sütunları sil
        columns_to_drop = ["UrunAdi", "Resim", "VaryasyonHepsiBuradaKodu", "ID"]
        merged_data_copy.drop(columns=columns_to_drop, inplace=True, errors='ignore')

        # Yenilenen değerleri teke düşür (benzersiz yap)
        merged_data_copy.drop_duplicates(inplace=True)

        # Yenilenmiş verileri bir dosyaya kaydet
        merged_data_copy.to_excel("birlesmis_veri_kopya.xlsx", index=False)




        # birlesmis_veri_kopya Excel dosyasını oku
        kopya_data = pd.read_excel("birlesmis_veri_kopya.xlsx")

        # birlesmis_veri Excel dosyasını oku
        veri_data = pd.read_excel("birlesmis_veri.xlsx")

        # Karşılık gelen verileri saklamak için bir liste oluştur
        results = []

        # birlesmis_veri_kopya'daki her bir ModelKodu için
        for model_kodu in kopya_data["ModelKodu"]:
            # birlesmis_veri'de ModelKodu'nu ara ve karşılık gelen verileri al
            match = veri_data[veri_data["ModelKodu"] == model_kodu]
            if not match.empty:
                # Karşılık gelen verileri results listesine ekle
                results.append(match[["ModelKodu", "UrunAdi", "Resim", "ID"]].values.tolist()[0])

        # Sonuçları bir DataFrame'e dönüştür
        result_df = pd.DataFrame(results, columns=["ModelKodu", "UrunAdi", "Resim", "ID"])

        # "Raf Ömrü" sütununu oluştur ve değerleri atayın
        result_df["Raf Ömrü (Ör: 12.12.2012-15.15.2013)"] = ["" for _ in range(len(result_df))]  # Örnek olarak hepsine "12 ay" atadım

        # DataFrame'i belirlenen parça sayısına göre bölelim
        parts = [result_df[i:i+len(result_df)//num_parts] for i in range(0, len(result_df), len(result_df)//num_parts)]

        # Her bir parçayı ayrı bir Excel dosyasına kaydedelim
        for i, part in enumerate(parts):
            with pd.ExcelWriter(f"Raf Ömrü {i+1}.xlsx", engine='xlsxwriter') as writer:
                # Sonuçları yaz
                part.to_excel(writer, index=False, sheet_name=f'Sheet1')
                # Excel dosyasının objesini al
                workbook = writer.book
                # DataFrame'in sütun başlıklarını al
                header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top',
                                                    'fg_color': '#D7E4BC', 'border': 1})
                # Başlıklara filtre özelliği ekle
                worksheet = writer.sheets['Sheet1']
                worksheet.autofilter(0, 0, part.shape[0], part.shape[1] - 1)  # Filtreyi ekle
                for col_num, value in enumerate(part.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                # Tüm sütunların genişliğini ayarla
                for i, col in enumerate(part.columns):
                    worksheet.set_column(i, i, 50)





        import os

        # Dosya adlarını tanımla
        dosya1 = "birlesmis_veri.xlsx"
        dosya2 = "birlesmis_veri_kopya.xlsx"

        # Dosyaları sil
        try:
            os.remove(dosya1)
            os.remove(dosya2)
        except FileNotFoundError:
            pass
        except Exception as e:
            print(f"Hata oluştu: {e}")


    except Exception as e:
        pass

    exit()


else:
    print("Geçersiz seçim.")
    exit()

#endregion

#region // Sadece Etiketli Ürünler mi Sorusu ve Ürün Listesi İndirme

# İndirilecek linkler
if etiket_secimi == "E":
    links = ["https://www.siparis.haydigiy.com/FaprikaXls/B0JC0W/1/"]
else:
    links = [
        "https://www.siparis.haydigiy.com/FaprikaXls/ZIMVGV/1/",
        "https://www.siparis.haydigiy.com/FaprikaXls/ZIMVGV/2/",
        "https://www.siparis.haydigiy.com/FaprikaXls/ZIMVGV/3/"
    ]

# Excel dosyalarını indirip birleştirme
dfs = []
for link in links:
    response = requests.get(link)
    if response.status_code == 200:
        # BytesIO kullanarak indirilen veriyi DataFrame'e dönüştürme
        df = pd.read_excel(BytesIO(response.content))
        
        # Belirli sütunu ve kullanıcının girdiği değeri içeren satırları seçme
        selected_rows = df[df[kolon_adi].astype(str).str.contains(re.escape(kullanici_input), case=False, na=False)]
        dfs.append(selected_rows)
    else:
        print(f"Hata: {response.status_code} - {link}")

# Tüm seçilen verileri birleştirme
if dfs:
    final_df = pd.concat(dfs, ignore_index=True)
else:
    print("Uygun veri bulunamadı.")


# Seçilen sütunu içeren satırları birleştirme
merged_df = pd.concat(dfs, ignore_index=True)

# Belirli başlıklar dışındaki sütunları silme
selected_columns = ["UrunAdi", "StokAdedi", "AlisFiyati", "SatisFiyati", "Kategori", "Resim", "AramaTerimleri", "MorhipoKodu", "VaryasyonMorhipoKodu", "HepsiBuradaKodu", "Marka", "N11Kodu"]
filtered_df = merged_df[selected_columns]

# Sonuç DataFrame'i tek bir Excel dosyasına yazma
filtered_df.to_excel("sonuc_excel.xlsx", index=False)

print(Fore.GREEN + "BAŞARILI - Ürün Listesi İndirme ve Sütun Ayarlamaları")

#endregion

#region // Ürünlerin Kategorilerini Belirleme ve Tesettür Ayarlaması

# Excel dosyasını okuma
df = pd.read_excel("sonuc_excel.xlsx")

# NaN değerleri boş string ile doldurma
df['Kategori'] = df['Kategori'].fillna("")

# Kategori sütunundan istenilen kısmı ayıklama fonksiyonu
def extract_category(text):
    if not isinstance(text, str):
        return None
    match = re.search(r'>\s*([^;]+)', text)
    if match:
        return match.group(1).strip()
    elif "TESETTÜR" in text:
        return "TESETTÜR"
    return None

# Yeni bir sütun oluşturup ayıklanan veriyi ekleme
df['Kategori'] = df['Kategori'].apply(extract_category)

# Yeni DataFrame'i bir Excel dosyasına kaydetme
df.to_excel("sonuc_excel.xlsx", index=False)

print(Fore.GREEN + "BAŞARILI - Ürünlerin Kategorilerini Belirleme ve Tesettür Ayarlaması")

#endregion

#region // UrunAdi Duzenleme Sütununu Oluşturma ve Sadece Ürün Kodlarını Bırakma

# sonuc_excel.xlsx dosyasını oku
sonuc_excel_file = "sonuc_excel.xlsx"
sonuc_df = pd.read_excel(sonuc_excel_file)

# Ürün kodunu ayıklamak için güncellenmiş bir fonksiyon tanımla
def extract_product_code(urun_adi):
    match = re.search(r' - (\d+)\.', urun_adi)  # " - " ve "." arasındaki sayıyı ayıkla
    return match.group(1) if match else None

# Yeni sütun oluştur ve her satır için ürün kodunu çek
sonuc_df['UrunAdi Duzenleme'] = sonuc_df['UrunAdi'].apply(extract_product_code)

# "UrunAdi Duzenleme" sütununu metin formatına çevir
sonuc_df['UrunAdi Duzenleme'] = sonuc_df['UrunAdi Duzenleme'].astype(str)

# Güncellenmiş DataFrame'i aynı Excel dosyasına kaydet
updated_excel_file = "sonuc_excel.xlsx"
sonuc_df.to_excel(updated_excel_file, index=False)

print(Fore.GREEN + "BAŞARILI - UrunAdi Duzenleme Sütununu Oluşturma ve Sadece Ürün Kodlarını Bırakma")

#endregion

#region // GMT ve SİTA Verilerini Çekme

# Google Sheet URL
google_sheet_url = "https://docs.google.com/spreadsheets/d/1aA5LhkQYgtwHLcKRV1mKl9Lb6VeOgUNIC9zy2kRagrs/gviz/tq?tqx=out:csv"

# Google Sheet'ten veriyi al ve Excel dosyasına kaydet
try:
    google_df = pd.read_csv(google_sheet_url)
    
    # "GMT Ürün Kodu" ve "SİTA Ürün Kodu" sütunlarındaki " - " ifadesinden sonrasını temizle
    google_df["GMT Ürün Kodu"] = google_df["GMT Ürün Kodu"].str.split(" - ").str[0]
    google_df["SİTA Ürün Kodu"] = google_df["SİTA Ürün Kodu"].str.split(" - ").str[0]
    
    # Sayıya çevirme işlemi ve hataları geçme
    google_df["GMT Ürün Kodu"] = pd.to_numeric(google_df["GMT Ürün Kodu"], errors='coerce')
    google_df["SİTA Ürün Kodu"] = pd.to_numeric(google_df["SİTA Ürün Kodu"], errors='coerce')

    # Hata nedeniyle NaN olan değerleri orijinal metin haline geri çevir
    google_df["GMT Ürün Kodu"] = google_df["GMT Ürün Kodu"].fillna(google_df["GMT Ürün Kodu"].astype(str))
    google_df["SİTA Ürün Kodu"] = google_df["SİTA Ürün Kodu"].fillna(google_df["SİTA Ürün Kodu"].astype(str))
    
    # Güncellenmiş DataFrame'i Excel dosyasına kaydet
    google_excel_file = "GMT ve SİTA.xlsx"
    google_df.to_excel(google_excel_file, index=False)
    
except requests.exceptions.RequestException as e:
    print(f"Request failed: {e}")

print(Fore.GREEN + "BAŞARILI - GMT ve SİTA Verilerine Erişme")

#endregion

#region // GMT ve SİTA Verilerini Ana Tabloya Çektirme (Etopla Yapma)

# sonuc_excel dosyasını oku
sonuc_excel_file = "sonuc_excel.xlsx"
sonuc_df = pd.read_excel(sonuc_excel_file)

# GMT ve SİTA dosyasını oku
gmt_sita_df = pd.read_excel("GMT ve SİTA.xlsx")

# İlk adım: 'UrunAdi' sütunuyla eşleşme ve GMT/SİTA Stok Adedi değerlerini al
def match_and_update(row):
    urun_adi = row['UrunAdi']
    
    # GMT Ürün Adı sütununda arama yap ve 'GMT Stok Adedi' değerini al
    gmt_row = gmt_sita_df[gmt_sita_df['GMT Ürün Adı'] == urun_adi]
    if not gmt_row.empty:
        row['GMT Stok Adedi'] = gmt_row['GMT Stok Adedi'].values[0]
    else:
        row['GMT Stok Adedi'] = None

    # SİTA Ürün Adı sütununda arama yap ve 'SİTA Stok Adedi' değerini al
    sista_row = gmt_sita_df[gmt_sita_df['SİTA Ürün Adı'] == urun_adi]
    if not sista_row.empty:
        row['SİTA Stok Adedi'] = sista_row['SİTA Stok Adedi'].values[0]
    else:
        row['SİTA Stok Adedi'] = None

    return row

# İlk adımı uygulayın
sonuc_df = sonuc_df.apply(match_and_update, axis=1)

# İkinci adım: 'UrunAdi Duzenleme' sütununa göre arama, sadece GMT Stok Adedi veya SİTA Stok Adedi boş ya da sıfırsa
def match_and_update_with_code(row):
    urun_kodu = row['UrunAdi Duzenleme']
    
    # Eğer GMT Stok Adedi boş ya da sıfırsa, 'GMT Ürün Kodu' ile arama yap
    if pd.isna(row['GMT Stok Adedi']) or row['GMT Stok Adedi'] == 0:
        gmt_code_row = gmt_sita_df[gmt_sita_df['GMT Ürün Kodu'] == urun_kodu]
        if not gmt_code_row.empty:
            gmt_etopla = gmt_code_row['GMT Stok Adedi'].values[0]
            row['GMT Stok Adedi'] = "GMT'de Var" if gmt_etopla > 0 else gmt_etopla
    
    # Eğer SİTA Stok Adedi boş ya da sıfırsa, 'SİTA Ürün Kodu' ile arama yap
    if pd.isna(row['SİTA Stok Adedi']) or row['SİTA Stok Adedi'] == 0:
        sista_code_row = gmt_sita_df[gmt_sita_df['SİTA Ürün Kodu'] == urun_kodu]
        if not sista_code_row.empty:
            sita_etopla = sista_code_row['SİTA Stok Adedi'].values[0]
            row['SİTA Stok Adedi'] = "SİTA'da Var" if sita_etopla > 0 else sita_etopla

    return row

# İkinci adımı uygulayın
sonuc_df = sonuc_df.apply(match_and_update_with_code, axis=1)

# Güncellenmiş DataFrame'i yeni bir Excel dosyasına kaydedin
updated_excel_file = "sonuc_excel.xlsx"
sonuc_df.to_excel(updated_excel_file, index=False)

print(Fore.GREEN + "BAŞARILI - GMT ve SİTA Verilerini Ana Tabloya Çektirme")

#endregion

#region // Stok Adedi Sütunu İçin Etopla Yapma ve Diğer Ayarlar

# "sonuc_excel.xlsx" Excel dosyasını oku
df_calisma_alani = pd.read_excel('sonuc_excel.xlsx')

# Aynı "UrunAdi" hücrelerinin "StokAdedi" sayılarını toplama
df_calisma_alani.loc[:, "StokAdedi"] = df_calisma_alani.groupby("UrunAdi")["StokAdedi"].transform("sum")

# "VaryasyonMorhipoKodu" sütununun adını değiştirme
df_calisma_alani = df_calisma_alani.rename(columns={"VaryasyonMorhipoKodu": "N11 & Zimmet"})

# Hesaplamalarda metinsel verileri sıfır olarak ele almak için sayısal değerlere dönüştürme
# Orijinal veri bozulmadan yalnızca matematiksel işlemler için geçici sütunlar kullanılıyor
gmt_numeric = pd.to_numeric(df_calisma_alani["GMT Stok Adedi"], errors="coerce").fillna(0)
sita_numeric = pd.to_numeric(df_calisma_alani["SİTA Stok Adedi"], errors="coerce").fillna(0)
stok_adedi_numeric = pd.to_numeric(df_calisma_alani["StokAdedi"], errors="coerce").fillna(0)
n11_zimmet_numeric = pd.to_numeric(df_calisma_alani["N11 & Zimmet"], errors="coerce").fillna(0)

# "Toplam Stok Adedi" sütunlarını oluşturma
df_calisma_alani["Stok Adedi Her Şey Dahil"] = stok_adedi_numeric + n11_zimmet_numeric + gmt_numeric + sita_numeric
df_calisma_alani["Stok Adedi Site ve Vega"] = stok_adedi_numeric + n11_zimmet_numeric

# Eksik değerleri sıfır ile doldurma (diğer sütunlar için)
df_calisma_alani['StokAdedi'].fillna(0, inplace=True)
df_calisma_alani['N11 & Zimmet'].fillna(0, inplace=True)
df_calisma_alani['GMT Stok Adedi'].fillna(0, inplace=True)
df_calisma_alani['SİTA Stok Adedi'].fillna(0, inplace=True)

# Güncellenmiş DataFrame'i yeni bir Excel dosyasına kaydet
updated_excel_file = "sonuc_excel.xlsx"
df_calisma_alani.to_excel(updated_excel_file, index=False)

print(Fore.GREEN + "BAŞARILI - Stok Adedi Sütunu İçin Etopla Yapma - Stok Adedi Her Şey Dahil ve Stok Adedi Site ve Vega Sütunlarını Oluşturma - Bazı Sütunların Adını Değiştirme")

#endregion

#region // MorhipoKodu Sütununun Adını Değiştirme ve Kaç Güne Biter Kısımlarını Hesaplama

# "MorhipoKodu" sütununun adını değiştirme /Komplo orduların
df_calisma_alani = df_calisma_alani.rename(columns={"MorhipoKodu": "Günlük Ortalama Satış Adedi"})
df_calisma_alani['Günlük Ortalama Satış Adedi'].fillna(0, inplace=True)

# "Kaç Güne Biter" sütununu oluşturma ve "Toplam Stok Adedi" sütunundaki verileri "Günlük Ortalama Satış Adedi" sütunundaki verilere bölme işlemi
df_calisma_alani["Kaç Güne Biter Her Şey Dahil"] = "Satış Adedi Yok"  # Varsayılan değer olarak "Satış Adedi Yok" atanır
df_calisma_alani["Kaç Güne Biter Site ve Vega"] = "Satış Adedi Yok"  # Varsayılan değer olarak "Satış Adedi Yok" atanır


non_zero_mask = df_calisma_alani["Günlük Ortalama Satış Adedi"] != 0
df_calisma_alani.loc[non_zero_mask, "Kaç Güne Biter Her Şey Dahil"] = round(df_calisma_alani["Stok Adedi Her Şey Dahil"] / df_calisma_alani["Günlük Ortalama Satış Adedi"])


non_zero_mask = df_calisma_alani["Günlük Ortalama Satış Adedi"] != 0
df_calisma_alani.loc[non_zero_mask, "Kaç Güne Biter Site ve Vega"] = round(df_calisma_alani["Stok Adedi Site ve Vega"] / df_calisma_alani["Günlük Ortalama Satış Adedi"])

print(Fore.GREEN + "BAŞARILI - MorhipoKodu Sütununun Adını Değiştirme ve Kaç Güne Biter Kısımlarını Hesaplama")

#endregion

#region // Görüntülenmenin Satışa Dönüş Oranını Hesaplama

# "Görüntülenmenin Satışa Dönüş Oranı" sütunu
df_calisma_alani["Görüntülenmenin Satışa Dönüş Oranı"] = "0"  # Varsayılan değer olarak "Satış Yok" atanır
df_calisma_alani = df_calisma_alani.rename(columns={"HepsiBuradaKodu": "Ortalama Görüntülenme Adedi"})
df_calisma_alani['Ortalama Görüntülenme Adedi'].fillna(0, inplace=True)
non_zero_mask = df_calisma_alani["Ortalama Görüntülenme Adedi"] != 0
df_calisma_alani.loc[non_zero_mask, "Görüntülenmenin Satışa Dönüş Oranı"] = round((df_calisma_alani["Günlük Ortalama Satış Adedi"] / df_calisma_alani["Ortalama Görüntülenme Adedi"]) * 100, 2)

# Değişiklikleri kaydetmek için dosyayı yeniden yaz
df_calisma_alani.to_excel("sonuc_excel.xlsx", index=False)

print(Fore.GREEN + "BAŞARILI - Görüntülenmenin Satışa Dönüş Oranını Hesaplama")

#endregion

#region // Satış Raporu Tarihini Düne Göre Ayarlama

# Excel dosyasının ismi ve konumu
filename = "Satış Raporu.xlsx"

# Dosyanın indirilme tarihini kontrol eden fonksiyon
def is_file_downloaded_today(file_path):
    if os.path.exists(file_path):
        # Dosyanın son değiştirilme tarihini al
        file_modification_time = os.path.getmtime(file_path)
        modification_date = datetime.fromtimestamp(file_modification_time).date()
        # Bugünün tarihi ile karşılaştır
        return modification_date == datetime.today().date()
    return False

# Eğer dosya bugün indirilmemişse Selenium işlemleri çalıştırılır
if not is_file_downloaded_today(filename):
    # ChromeOptions ile gizli mod seçeneğini ayarla
    options = Options()
    options.add_argument("--headless")  # Tarayıcıyı ekranda göstermemek için
    options.add_argument("--incognito")  # Gizli mod için bu parametre eklenir
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    # ChromeDriver'ı en son sürümüyle otomatik olarak indirip kullan
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    login_url = "https://www.siparis.haydigiy.com/kullanici-giris/?ReturnUrl=%2Fadmin"
    driver.get(login_url)

    # Giriş bilgilerini doldurma
    email_input = driver.find_element("id", "EmailOrPhone")
    email_input.send_keys("mustafa_kod@haydigiy.com")

    password_input = driver.find_element("id", "Password")
    password_input.send_keys("123456")
    password_input.send_keys(Keys.RETURN)

    # Belirttiğiniz sayfaya yönlendirme
    desired_page_url = "https://www.siparis.haydigiy.com/admin/exportorder/edit/154/"
    driver.get(desired_page_url)

    # Dünün tarihini al
    yesterday = datetime.now() - timedelta(days=1)
    formatted_date = yesterday.strftime("%d.%m.%Y")

    # EndDate alanını bulma ve tarih girişini yapma
    end_date_input = driver.find_element("id", "EndDate")
    end_date_input.clear()  # Eğer mevcut bir değer varsa temizleyin
    end_date_input.send_keys(formatted_date)

    # StartDate alanını bulma ve tarih girişini yapma
    start_date_input = driver.find_element("id", "StartDate")
    start_date_input.clear()  # Eğer mevcut bir değer varsa temizleyin
    start_date_input.send_keys(formatted_date)

    # Kaydet butonunu bulma ve tıklama
    save_button = driver.find_element("css selector", 'button.btn.btn-primary[name="save"]')
    save_button.click()

    # Selenium işlemleri tamamlandıktan sonra tarayıcıyı kapatın
    driver.quit()

print(Fore.GREEN + "BAŞARILI - Sitedeki Satış Raporu Çıktısının Tarihini Dün Olarak Ayarlama")

#endregion

#region // Satış Raporunu İndirme

# Excel dosyasının indirileceği URL
url = "https://www.siparis.haydigiy.com/FaprikaOrderXls/GZPCKE/1/"
filename = "Satış Raporu.xlsx"

# Dosyanın indirilme tarihini kontrol etmek için fonksiyon
def is_file_downloaded_today(file_path):
    if os.path.exists(file_path):
        # Dosyanın son değiştirilme tarihini al
        file_modification_time = os.path.getmtime(file_path)
        modification_date = datetime.fromtimestamp(file_modification_time).date()
        # Bugünün tarihi ile karşılaştır
        return modification_date == datetime.today().date()
    return False

# Dosya bugün indirilmemişse veya yoksa yeniden indir
if not is_file_downloaded_today(filename):
    # Eğer dosya varsa sil
    if os.path.exists(filename):
        os.remove(filename)
    # Dosyayı indir ve kaydet
    response = requests.get(url)
    with open(filename, "wb") as file:
        file.write(response.content)

# Excel dosyasını oku
df = pd.read_excel(filename)

# Tutulacak sütunlar
columns_to_keep = ["UrunAdi", "Adet", "ToplamFiyat"]

# Diğer sütunları silmek
df = df[columns_to_keep]

# Düzenlenmiş dosyayı aynı adla kaydet
df.to_excel(filename, index=False)

print(Fore.GREEN + "BAŞARILI - Satış Raporu İndirme")

#endregion

#region // Adet Sütununu Sayıya Çevirme

def clean_adet(data):
    # "Adet" sütunundaki tüm verilerin virgül karakterinden sonrasını temizle
    data['Adet'] = data['Adet'].astype(str).apply(lambda x: x.split(',')[0])

if __name__ == "__main__":
    # Mevcut Excel dosyasını oku
    file_path = "Satış Raporu.xlsx"
    combined_data = pd.read_excel(file_path, engine="openpyxl")

    # "Adet" sütunundaki verilerin virgül karakterinden sonrasını temizle
    clean_adet(combined_data)

    # Güncellenmiş veriyi aynı dosyaya üstüne kaydet
    combined_data.to_excel(file_path, index=False, engine='openpyxl')
    
   
def convert_adet_to_numeric(data):
    # "Adet" sütunundaki tüm verileri sayıya dönüştür
    data['Adet'] = pd.to_numeric(data['Adet'], errors='coerce')  # 'coerce' ile hatalı değerleri NaN olarak işaretle

if __name__ == "__main__":
    # Mevcut Excel dosyasını oku
    file_path = "Satış Raporu.xlsx"
    combined_data = pd.read_excel(file_path, engine="openpyxl")

    # "Adet" sütunundaki verileri sayıya dönüştür
    convert_adet_to_numeric(combined_data)

    # Güncellenmiş veriyi aynı dosyaya üstüne kaydet
    combined_data.to_excel(file_path, index=False, engine='openpyxl')

print(Fore.GREEN + "BAŞARILI - Satış Raporu Düzenleme 1")

#endregion

#region // ToplamFiyat Sütununu Sayıya Çevirme

def clean_adet(data):
    # "Adet" sütunundaki tüm verilerin virgül karakterinden sonrasını temizle
    data['ToplamFiyat'] = data['ToplamFiyat'].astype(str).apply(lambda x: x.split(',')[0])

if __name__ == "__main__":
    # Mevcut Excel dosyasını oku
    file_path = "Satış Raporu.xlsx"
    combined_data = pd.read_excel(file_path, engine="openpyxl")

    # "Adet" sütunundaki verilerin virgül karakterinden sonrasını temizle
    clean_adet(combined_data)

    # Güncellenmiş veriyi aynı dosyaya üstüne kaydet
    combined_data.to_excel(file_path, index=False, engine='openpyxl')
    
   
def convert_adet_to_numeric(data):
    # "Adet" sütunundaki tüm verileri sayıya dönüştür
    data['ToplamFiyat'] = pd.to_numeric(data['ToplamFiyat'], errors='coerce')  # 'coerce' ile hatalı değerleri NaN olarak işaretle

if __name__ == "__main__":
    # Mevcut Excel dosyasını oku
    file_path = "Satış Raporu.xlsx"
    combined_data = pd.read_excel(file_path, engine="openpyxl")

    # "Adet" sütunundaki verileri sayıya dönüştür
    convert_adet_to_numeric(combined_data)

    # Güncellenmiş veriyi aynı dosyaya üstüne kaydet
    combined_data.to_excel(file_path, index=False, engine='openpyxl')

print(Fore.GREEN + "BAŞARILI - Satış Raporu Düzenleme 2")

#endregion

#region // Adet ve ToplamFiyat Sütununa ETOPLA yapma

# Excel dosyasını tekrar okumak
df = pd.read_excel("Satış Raporu.xlsx")

# UrunAdi sütununa göre gruplandırma ve Adet ile ToplamFiyat sütunlarındaki verileri toplama
df_grouped = df.groupby('UrunAdi', as_index=False).agg({
    'Adet': 'sum',
    'ToplamFiyat': 'sum'
})

# Düzenlenmiş dosyayı aynı adla kaydetmek
df_grouped.to_excel("Satış Raporu.xlsx", index=False)

print(Fore.GREEN + "BAŞARILI - Satış Raporu Düzenleme 3")

#endregion

#region // Ana Listeye Veriyi Çektirme

# Excel dosyalarını oku
satis_raporu_df = pd.read_excel("Satış Raporu.xlsx")
one_cikanlar_df = pd.read_excel("sonuc_excel.xlsx")

# Öne Çıkanlar Excel'ine Satış Raporu'ndan Adet ve ToplamFiyat sütunlarını eklemek için merge işlemi yapalım
merged_df = one_cikanlar_df.merge(
    satis_raporu_df[['UrunAdi', 'Adet']],
    on='UrunAdi',
    how='left'
)

# Sütun adını değiştir
merged_df.rename(columns={'Adet': 'Dünün Satış Adedi'}, inplace=True)

# Birleştirilmiş veriyi Öne Çıkanlar Excel dosyasına kaydedelim
merged_df.to_excel("sonuc_excel.xlsx", index=False)

print(Fore.GREEN + "BAŞARILI - Ana Tabloya Satış Adetlerini Çektime")

#endregion

#region // Sütunların Sırasını Değiştirme - Bazı Sütunların Adını Değiştirme

# Excel dosyasını oku
df_calisma_alani = pd.read_excel("sonuc_excel.xlsx")

# "StokAdedi" sütununun adını değiştirme
df_calisma_alani = df_calisma_alani.rename(columns={"StokAdedi": "İnstagram Stok Adedi"})

# Sütun sıralamasını ayarlama
column_order = ["UrunAdi", "İnstagram Stok Adedi", "Stok Adedi Her Şey Dahil", "Stok Adedi Site ve Vega", 
                "Günlük Ortalama Satış Adedi", "Dünün Satış Adedi", "Ortalama Görüntülenme Adedi", "Görüntülenmenin Satışa Dönüş Oranı", 
                "Kaç Güne Biter Her Şey Dahil", "Kaç Güne Biter Site ve Vega", "AlisFiyati", "SatisFiyati", 
                "AramaTerimleri", "Resim", "Kategori", "GMT Stok Adedi", "SİTA Stok Adedi", "Marka", "N11Kodu"]
df_calisma_alani = df_calisma_alani[column_order]

print(Fore.GREEN + "BAŞARILI - Sütunların Sırasını Değiştirme - Bazı Sütunların Adını Değiştirme")

#endregion

#region // Yenilenen Değerleri Kaldırma 

# Tekrarlanan satırları silme
df_calisma_alani = df_calisma_alani.drop_duplicates(subset=["UrunAdi"])

print(Fore.GREEN + "BAŞARILI - Yenilenen Değerleri Kaldırma")

#endregion

#region // Resim Sütunu İçin .jpeg'den Sonrasını Kladırma ve Devamına .jpeg Ekleme

# "Resim" sütunundaki ".jpeg" ifadesinden sonrasını temizleme ve ".jpeg" ekleme
df_calisma_alani["Resim"] = df_calisma_alani["Resim"].str.replace(r"\.jpeg.*$", "", regex=True) + ".jpeg"

# Resim bağlantılarını bir listeye kaydet
links = df_calisma_alani["Resim"].tolist()

# NaN değerlerini 0 ile değiştirme
df_calisma_alani = df_calisma_alani.fillna(0)

# inf değerlerini 0 ile değiştirme
df_calisma_alani.replace([float('inf'), float('-inf')], 0, inplace=True)

print(Fore.GREEN + "BAŞARILI - Resim Sütununu Ayarlama ve Köprü Verme")

#endregion

#region // AramaTerimleri Sütunundaki Tarihleri Ayıklama

# "AramaTerimleri" sütunundaki tarihleri temizle
date_pattern = r'(\d{1,2}\.\d{1,2}\.\d{4})'
df_calisma_alani['AramaTerimleri'] = df_calisma_alani['AramaTerimleri'].apply(lambda x: re.search(date_pattern, str(x)).group(1) if re.search(date_pattern, str(x)) else None)

print(Fore.GREEN + "BAŞARILI - AramaTerimleri Sütunundaki Tarihleri Ayıklama")

#endregion

#region // Bazı Sütunların Adını Güncelleme

# "AramaTerimleri" sütununun adını "Resim Yüklenme Tarihi" olarak değiştirme
df_calisma_alani.rename(columns={"AramaTerimleri": "Resim Yüklenme Tarihi"}, inplace=True)
df_calisma_alani.rename(columns={"AlisFiyati": "Alış Fiyatı"}, inplace=True)
df_calisma_alani.rename(columns={"SatisFiyati": "Satış Fiyatı"}, inplace=True)
df_calisma_alani.rename(columns={"UrunAdi": "Ürün Adı"}, inplace=True)
df_calisma_alani.rename(columns={"N11Kodu": "Mevsim"}, inplace=True)

print(Fore.GREEN + "BAŞARILI - Sütun İsimleri Güncelleme")

#endregion

#region // Resim Sütunuyla Alakalı Bir İşlem

# "Resim" sütununu DataFrame'den kaldır
df_calisma_alani.drop(columns=["Resim"], inplace=True)

print(Fore.GREEN + "BAŞARILI - Resim Sütununu DataFrameden Kaldırma")

#endregion

#region // Sütunların Biçim Ayarları ve Bazı Ayarlamalar

# Güncellenmiş DataFrame'i aynı Excel dosyasına yaz
with pd.ExcelWriter('sonuc_excel.xlsx', engine='xlsxwriter') as writer:
    df_calisma_alani.to_excel(writer, index=False, sheet_name='Sheet1')

    # ExcelWriter objesinden workbook ve worksheet'e eriş
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # İlk sütun (Ürün Adı) için uygun genişlik ayarlama
    max_col_width = max(df_calisma_alani["Ürün Adı"].astype(str).apply(len).max(), len("Ürün Adı")) + 2
    worksheet.set_column(0, 0, max_col_width)

    # Belirli sütunlar için genişliği 10 piksel olarak ayarlama
    narrow_columns = ["Alış Fiyatı", "Satış Fiyatı", "GMT Stok Adedi", "SİTA Stok Adedi"]
    for col_name in narrow_columns:
        col_idx = df_calisma_alani.columns.get_loc(col_name)
        worksheet.set_column(col_idx, col_idx, 10)

    # Diğer tüm sütunların genişliğini 15 piksel olarak ayarla
    for i in range(1, len(df_calisma_alani.columns)):
        if df_calisma_alani.columns[i] not in narrow_columns:
            worksheet.set_column(i, i, 15)

    # Başlık satırının yüksekliğini 50 piksel olarak ayarla
    worksheet.set_row(0, 50)

    # Başlık satırını dondur
    worksheet.freeze_panes(1, 0)

    # Başlık için hücre biçimlendirme ayarlarını yap
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#D3D3D3',
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
        'border': 1
    })

    # Para birimi formatı tanımlama (orta hizalı)
    currency_format = workbook.add_format({'num_format': '#,##0.00₺', 'align': 'center', 'valign': 'vcenter'})
    shaded_currency_format = workbook.add_format({'bg_color': '#D9D9D9', 'num_format': '#,##0.00₺', 'align': 'center', 'valign': 'vcenter'})

    # Var içeren hücreler için özel renklendirme
    var_format = workbook.add_format({'bg_color': '#ffb994'})

    # Başlık hücrelerini yaz ve biçimlendir
    for col_num, value in enumerate(df_calisma_alani.columns.values):
        worksheet.write(0, col_num, value, header_format)

    # Veriler için hücre biçimlendirme ayarlarını yap
    center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
    left_align_format = workbook.add_format({'align': 'left', 'valign': 'vcenter'})
    shaded_format = workbook.add_format({'bg_color': '#D9D9D9', 'align': 'center', 'valign': 'vcenter'})
    shaded_left_align_format = workbook.add_format({'bg_color': '#D9D9D9', 'align': 'left', 'valign': 'vcenter'})

    for row_num, row in enumerate(df_calisma_alani.itertuples(), start=1):
        for col_num, value in enumerate(row[1:]):  # row[0] index olduğu için atlanıyor
            col_name = df_calisma_alani.columns[col_num]
            try:
                if col_num == 0:  # Ürün Adı sütunu
                    link = links[row_num - 1]  # Resim sütunundaki bağlantıyı Ürün Adı'ya ekliyoruz
                    if isinstance(link, str) and link.startswith("http"):  # Link geçerli mi?
                        if row_num % 2 == 1:
                            worksheet.write_url(row_num, col_num, link, string=value, cell_format=shaded_left_align_format)
                        else:
                            worksheet.write_url(row_num, col_num, link, string=value, cell_format=left_align_format)
                    else:  # Link geçerli değilse düz metin olarak yaz
                        if row_num % 2 == 1:
                            worksheet.write(row_num, col_num, value, shaded_left_align_format)
                        else:
                            worksheet.write(row_num, col_num, value, left_align_format)
                elif col_name in ["Alış Fiyatı", "Satış Fiyatı"]:
                    # Para birimi formatı uygulama, alternatif satır renklendirme ile ve orta hizalı
                    if row_num % 2 == 1:
                        worksheet.write(row_num, col_num, value, shaded_currency_format)
                    else:
                        worksheet.write(row_num, col_num, value, currency_format)
                elif col_name in ["GMT Stok Adedi", "SİTA Stok Adedi"] and "Var" in str(value):
                    worksheet.write(row_num, col_num, value, var_format)
                else:
                    if row_num % 2 == 1:
                        worksheet.write(row_num, col_num, value, shaded_format)
                    else:
                        worksheet.write(row_num, col_num, value, center_format)
            except Exception as e:
                # Hata durumunda devam et
                print(f"Hata: Satır {row_num}, Sütun {col_num}, Değer: {value}, Hata Mesajı: {e}")
                continue

    # Başlıklara filtre ekleme
    worksheet.autofilter(0, 0, 0, len(df_calisma_alani.columns) - 1)

    # Sayfanın yakınlaştırma oranını %90 olarak ayarla
    worksheet.set_zoom(90)

# Dosyanın adını değiştirme
excel_file_name = "sonuc_excel.xlsx"
new_excel_file_name = "Nirvana.xlsx"
os.rename(excel_file_name, new_excel_file_name)

print(Fore.GREEN + "BAŞARILI - Sütunların Biçim Ayarları ve Diğer Ayarlamalar")

#endregion

#region // Gereksiz Excel Dosyalarını Silme

# Eski dosyaları silme
dosyalar = ["GMT ve SİTA.xlsx", "Satış Raporu.xlsx"]

for dosya in dosyalar:
    if os.path.exists(dosya):
        os.remove(dosya)

print(Fore.GREEN + "BAŞARILI - Gereksiz Excel Dosyalarını Silme")

#endregion

#region // Sütunlara Açıklama Ekleme

# Excel dosyasını yükle
dosya_yolu = "Nirvana.xlsx"
workbook = load_workbook(dosya_yolu)
sheet = workbook.active

# Sütun başlıkları ve açıklama metinleri
columns_with_comments = {
    "İnstagram Stok Adedi": "Ürünün sitedeki satışa açık stok adedini belirtir",
    "Stok Adedi Her Şey Dahil": "Ürünün Instagram - STAD Depo - Zimmet Depo - GMT - SİTA kısımlarındaki toplam stok adedini belirtir",
    "Stok Adedi Site ve Vega": "Ürünün Instagram - STAD Depo - Zimmet Depo kısımlarındaki toplam stok adedini belirtir",
    "Günlük Ortalama Satış Adedi": "Ürünün son 1 haftaya göre toplam satış adedini son 1 haftaya göre kaç gündür aktif satışta olduğu güne böler ve haftanın ortalama satış adedini tespit eder",
    "Dünün Satış Adedi": "Ürünün dün sattığı adedi belirtir",
    "Ortalama Görüntülenme Adedi": "Ürünün son 1 haftaya göre toplam görüntülenme adedini son 1 haftaya göre kaç gündür aktif satışta olduğu güne böler ve haftanın ortalama görüntülenme adedini tespit eder",
    "Görüntülenmenin Satışa Dönüş Oranı": "Ürünün ortalama görüntülenme adedini, ortalama satış adedine bölerek görüntülenmenin ne kadar satışa dönüştüğünü belirtir",
    "Kaç Güne Biter Her Şey Dahil": "Ürünün Instagram - STAD Depo - Zimmet Depo - GMT - SİTA kısımlarındaki toplam stok adetlerinin ortalama satış adedine göre kaç günde biteceğini belirtir",
    "Kaç Güne Biter Site ve Vega": "Ürünün Instagram - STAD Depo - Zimmet Depo kısımlarındaki toplam stok adetlerinin ortalama satış adedine göre kaç günde biteceğini belirtir",
    "Alış Fiyatı": "Ürünün site üzerindeki güncel alış fiyatını belirtir",
    "Satış Fiyatı": "Ürünün site üzerindeki güncel satış fiyatını belirtir",
    "Resim Yüklenme Tarihi": "Ürünün resminin yüklenip satışa açıldığı tarihi belirtir",
    "Kategori": "Ürünün ana kategorisini belirtir",
    "GMT Stok Adedi": "Ürünün GMT üzerinde kalan olarak ne kadar stok adedi olduğunu belirtir",
    "SİTA Stok Adedi": "Ürünün SİTA üzerinde ne kadar stok adedi olduğunu belirtir"
}

# Başlık hücrelerini bul ve açıklama ekle
for cell in sheet[1]:  # 1. satırdaki tüm hücreleri kontrol eder
    if cell.value in columns_with_comments:
        # Yükseklik ve genişlik %100x%100 olacak şekilde açıklama oluştur
        comment = Comment(columns_with_comments[cell.value], "Açıklama", width=200, height=150)
        cell.comment = comment

# Değişiklikleri kaydet
workbook.save(dosya_yolu)

# Workbook nesnesini serbest bırak ve önbelleği temizle
del workbook
gc.collect()

print(Fore.GREEN + "BAŞARILI - Sütunlara Açıklama Ekleme")

#endregion

#region // Sigara Ürünleri Markadan Tespit Etme

# Excel dosyasını yükle
file_path = "Nirvana.xlsx"
workbook = openpyxl.load_workbook(file_path)
sheet = workbook["Sheet1"]

# "Ürün Adı" ve "Marka" sütunlarının indekslerini bul
urun_adi_column = None
marka_column = None

for col_index, column in enumerate(sheet[1], start=1):
    if column.value == "Ürün Adı":
        urun_adi_column = col_index
    elif column.value == "Marka":
        marka_column = col_index

# Hata kontrolü: Eğer "Ürün Adı" veya "Marka" sütunu bulunamazsa
if urun_adi_column is None or marka_column is None:
    raise ValueError("'Ürün Adı' veya 'Marka' sütunu bulunamadı.")

# Verileri kontrol edip hücre rengini değiştirme
for row in sheet.iter_rows(min_row=2):  # Başlık satırını atla
    urun_adi_cell = row[urun_adi_column - 1]
    marka_cell = row[marka_column - 1]

    if marka_cell.value and isinstance(marka_cell.value, str) and "Sigara Ürün" in marka_cell.value:
        urun_adi_cell.fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")  # Açık mavi renk

# "Marka" sütununu sil
sheet.delete_cols(marka_column)

# Değişiklikleri kaydet
workbook.save(file_path)

print(Fore.GREEN + "BAŞARILI - Sigara Ürünleri Markadan Tespit Etme")

#endregion


#region // Kopya Sayfa Oluşturma


# Nirvana.xlsx dosyasını yükle
dosya_adi = "Nirvana.xlsx"
workbook = openpyxl.load_workbook(dosya_adi)

# Sheet1 sayfasını kopyala
if "Sheet1" in workbook.sheetnames:
    sheet1 = workbook["Sheet1"]
    sheet_copy = workbook.copy_worksheet(sheet1)
    sheet_copy.title = "Sheet1_Copy"  # Yeni sayfa adı

# Dosyayı kaydet
workbook.save(dosya_adi)

print(Fore.GREEN + "BAŞARILI - Kopya Sayfa Oluşturma")

#endregion

#region // Kar Yüzdesi Sütununu Hesaplama

# Excel dosyasını yükle
dosya_adi = "Nirvana.xlsx"
workbook = openpyxl.load_workbook(dosya_adi)

# Kopya sayfayı seç
kopya_sayfa_adi = "Sheet1_Copy"
if kopya_sayfa_adi in workbook.sheetnames:
    sheet = workbook[kopya_sayfa_adi]

    # Başlıkları kontrol et ve "Stok Adedi Her Şey Dahil" sütununu bul
    basliklar = {cell.value: cell.column for cell in sheet[1] if cell.value}
    stok_adedi_kolon = basliklar.get("Stok Adedi Her Şey Dahil")

    if stok_adedi_kolon:
        # Yeni sütun indeksi (kopyalanan sütunun yanına eklenecek)
        yeni_sutun_index = stok_adedi_kolon + 1

        # "Stok Adedi Her Şey Dahil" sütununu biçimleriyle birlikte kopyala
        for row in range(1, sheet.max_row + 1):
            eski_hucre = sheet.cell(row=row, column=stok_adedi_kolon)
            yeni_hucre = sheet.cell(row=row, column=yeni_sutun_index)

            # Veriyi ve biçimlendirmeyi kopyala
            yeni_hucre.value = eski_hucre.value
            if eski_hucre.has_style:
                yeni_hucre._style = copy(eski_hucre._style)

        # Yeni sütuna başlık ekle
        sheet.cell(row=1, column=yeni_sutun_index).value = "Kar Yüzdesi"

        # Gerekli sütunların indekslerini belirle
        satis_fiyati_kolon = basliklar.get("Satış Fiyatı")
        alis_fiyati_kolon = basliklar.get("Alış Fiyatı")

        if satis_fiyati_kolon and alis_fiyati_kolon:
            # Kar yüzdesi hesaplamasını yap ve yüzde formatını uygula
            for row in range(2, sheet.max_row + 1):
                satis_fiyati = sheet.cell(row=row, column=satis_fiyati_kolon).value
                alis_fiyati = sheet.cell(row=row, column=alis_fiyati_kolon).value

                yeni_hucre = sheet.cell(row=row, column=yeni_sutun_index)

                if satis_fiyati and alis_fiyati:  # Boş hücreleri kontrol et
                    try:
                        kar_yuzdesi = (satis_fiyati - alis_fiyati) / satis_fiyati
                        yeni_hucre.value = kar_yuzdesi
                        yeni_hucre.number_format = "0.00%"  # Yüzde formatı
                    except ZeroDivisionError:
                        yeni_hucre.value = None

# Değişiklikleri kaydet
workbook.save(dosya_adi)

print(Fore.GREEN + "BAŞARILI - Kar Yüzdesi Sütununu Hesaplama")

#endregion

#region // Liste Fiyatını Hesaplama

# Excel dosyasını yükle
dosya_adi = "Nirvana.xlsx"
sheet_adi = "Sheet1_Copy"
workbook = openpyxl.load_workbook(dosya_adi)
if sheet_adi in workbook.sheetnames:
    sheet = workbook[sheet_adi]

    # Başlıkları ve kolon indekslerini belirle
    basliklar = {cell.value: cell.column for cell in sheet[1] if cell.value}
    alis_fiyati_kolon = basliklar.get("Alış Fiyatı")
    kategori_kolon = basliklar.get("Kategori")
    yeni_kolon_index = sheet.max_column + 1

    # Yeni sütuna başlık ekle
    sheet.cell(row=1, column=yeni_kolon_index).value = "ListeFiyati2"

    # Liste fiyatını hesaplayarak yeni sütuna ekle
    for row in range(2, sheet.max_row + 1):
        alis_fiyati = sheet.cell(row=row, column=alis_fiyati_kolon).value
        kategori = sheet.cell(row=row, column=kategori_kolon).value

        if alis_fiyati is not None:
            # Liste fiyatını hesaplama
            if 0 <= alis_fiyati <= 24.99:
                result = alis_fiyati + 10
            elif 25 <= alis_fiyati <= 39.99:
                result = alis_fiyati + 13
            elif 40 <= alis_fiyati <= 59.99:
                result = alis_fiyati + 17
            elif 60 <= alis_fiyati <= 200.99:
                result = alis_fiyati * 1.30
            elif alis_fiyati >= 201:
                result = alis_fiyati * 1.25
            else:
                result = alis_fiyati

            # KDV hesaplama
            if isinstance(kategori, str) and any(category in kategori for category in ["Parfüm", "Gözlük", "Saat", "Kolye", "Küpe", "Bileklik", "Bilezik"]):
                result *= 1.20
            else:
                result *= 1.10

            # Sonuç ekle
            sheet.cell(row=row, column=yeni_kolon_index).value = result

# Değişiklikleri kaydet
workbook.save(dosya_adi)

print(Fore.GREEN + "BAŞARILI - Liste Fiyatı Hesaplama")

#endregion

#region // Satış Fiyatı Liste Fiyatının Altındaysa Alış Fiyatını Kırmızı Yapma

# Excel dosyasını yükle
dosya_adi = "Nirvana.xlsx"
sheet_adi = "Sheet1_Copy"
workbook = openpyxl.load_workbook(dosya_adi)
if sheet_adi in workbook.sheetnames:
    sheet = workbook[sheet_adi]

    # Başlıkları ve kolon indekslerini belirle
    basliklar = {cell.value: cell.column for cell in sheet[1] if cell.value}
    liste_fiyati2_kolon = basliklar.get("ListeFiyati2")
    satis_fiyati_kolon = basliklar.get("Satış Fiyatı")
    alis_fiyati_kolon = basliklar.get("Alış Fiyatı")

    if liste_fiyati2_kolon and satis_fiyati_kolon and alis_fiyati_kolon:
        for row in range(2, sheet.max_row + 1):
            liste_fiyati2 = sheet.cell(row=row, column=liste_fiyati2_kolon).value
            satis_fiyati = sheet.cell(row=row, column=satis_fiyati_kolon).value
            alis_fiyati_hucre = sheet.cell(row=row, column=alis_fiyati_kolon)

            if liste_fiyati2 is not None and satis_fiyati is not None:
                fark = liste_fiyati2 - satis_fiyati  # ListeFiyati2 - Satış Fiyatı farkı

                # Fark 7'den büyükse yazı rengini kırmızı yap
                if fark > 7:
                    alis_fiyati_hucre.font = Font(color="FF0000")  # Kırmızı renk
                else:
                    # Eğer renk değiştirilmemesi gerekiyorsa hiçbir işlem yapılmaz
                    pass

# Değişiklikleri kaydet
workbook.save(dosya_adi)

print(Fore.GREEN + "BAŞARILI - Satış Fiyatı Liste Fiyatının Altındaysa Alış Fiyatını Kırmızı Yapma")

#endregion

#region // ListeFiyati2 Sütununu Silme

# Excel dosyasını yükle
dosya_adi = "Nirvana.xlsx"
sheet_adi = "Sheet1_Copy"
workbook = openpyxl.load_workbook(dosya_adi)

# Kopya sayfayı seç
if sheet_adi in workbook.sheetnames:
    sheet = workbook[sheet_adi]

    # "ListeFiyati2" sütununu bul ve sil
    basliklar = {cell.value: cell.column for cell in sheet[1] if cell.value}
    liste_fiyati2_kolon = basliklar.get("ListeFiyati2")

    if liste_fiyati2_kolon:
        sheet.delete_cols(liste_fiyati2_kolon)

# Değişiklikleri kaydet
workbook.save(dosya_adi)

print(Fore.GREEN + "BAŞARILI - ListeFiyati2 Sütununu Silme")

#endregion

#region // Belirli Sütunları Silme

# Nirvana.xlsx dosyasını yükle
dosya_adi = "Nirvana.xlsx"
workbook = openpyxl.load_workbook(dosya_adi)

# Kopya sayfayı seç
kopya_sayfa_adi = "Sheet1_Copy"
if kopya_sayfa_adi in workbook.sheetnames:
    sheet = workbook[kopya_sayfa_adi]

    # Silinecek sütunların adları
    silinecek_sutunlar = [
        "İnstagram Stok Adedi",
        "Stok Adedi Site ve Vega",
        "Ortalama Görüntülenme Adedi",
        "Kaç Güne Biter Site ve Vega",
        "Satış Fiyatı"
    ]

    # Başlıkları oku ve sütunları belirle
    basliklar = {cell.value: cell.column for cell in sheet[1] if cell.value is not None}

    # Sütunları tersten sil (indeks kaymasını önlemek için)
    for sutun_adi in reversed(silinecek_sutunlar):
        if sutun_adi in basliklar:
            sheet.delete_cols(basliklar[sutun_adi])

# Değişiklikleri kaydet
workbook.save(dosya_adi)

print(Fore.GREEN + "BAŞARILI - Belirli Sütunları Silme")

#endregion

#region // Sütunları Gizleme

# Nirvana.xlsx dosyasını tekrar yükle
workbook = openpyxl.load_workbook(dosya_adi)

# Kopya sayfayı seç
if kopya_sayfa_adi in workbook.sheetnames:
    sheet = workbook[kopya_sayfa_adi]

    # Gizlenecek sütunların adları
    gizlenecek_sutunlar = [
        "Resim Yüklenme Tarihi",
        "Kategori",
        "GMT Stok Adedi",
        "SİTA Stok Adedi"
    ]

    # Başlıkları oku ve sütunları belirle
    basliklar = {cell.value: cell.column for cell in sheet[1] if cell.value is not None}

    # Gizlenecek sütunları gizle
    for sutun_adi in gizlenecek_sutunlar:
        if sutun_adi in basliklar:
            sheet.column_dimensions[openpyxl.utils.get_column_letter(basliklar[sutun_adi])].hidden = True

# Değişiklikleri kaydet
workbook.save(dosya_adi)

print(Fore.GREEN + "BAŞARILI - Sütunları Gizleme")

#endregion

#region // Kar Yüzdesi Sütununu Görünür Yapma

# Nirvana.xlsx dosyasını yükle
dosya_adi = "Nirvana.xlsx"
workbook = openpyxl.load_workbook(dosya_adi)

# Kopya sayfayı seç
kopya_sayfa_adi = "Sheet1_Copy"
if kopya_sayfa_adi in workbook.sheetnames:
    sheet = workbook[kopya_sayfa_adi]

    # Başlıkları kontrol et ve "Kar Yüzdesi" sütununu bul
    basliklar = {cell.value: cell.column for cell in sheet[1] if cell.value is not None}
    if "Kar Yüzdesi" in basliklar:
        kar_yuzdesi_kolon = basliklar["Kar Yüzdesi"]

        # Sütunu görünür yap
        sheet.column_dimensions[openpyxl.utils.get_column_letter(kar_yuzdesi_kolon)].hidden = False

# Değişiklikleri kaydet
workbook.save(dosya_adi)

print(Fore.GREEN + "BAŞARILI - Kar Yüzdesi Sütununu Görünür Yapma")

#endregion

#region // Sütunlara Filtreleme Özelliği Ekleme

# Nirvana.xlsx dosyasını yükle
dosya_adi = "Nirvana.xlsx"
sheet_adi = "Sheet1_Copy"
workbook = openpyxl.load_workbook(dosya_adi)

# Belirtilen sayfayı seç
if sheet_adi in workbook.sheetnames:
    sheet = workbook[sheet_adi]

    # Sütun filtreleme özelliği ekleme
    max_row = sheet.max_row
    max_col = sheet.max_column
    filter_ref = f"A1:{openpyxl.utils.get_column_letter(max_col)}{max_row}"

    # AutoFilter özelliğini ekle
    sheet.auto_filter.ref = filter_ref

# Değişiklikleri kaydet
workbook.save(dosya_adi)

print(Fore.GREEN + "BAŞARILI - Sütunlara Filtreleme Özelliği Ekleme")

#endregion

