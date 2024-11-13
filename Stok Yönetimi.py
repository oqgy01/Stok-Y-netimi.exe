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

        from openpyxl.utils import get_column_letter
        import openpyxl

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




























# İndirilecek linkler
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


# Seçilen sütunu içeren satırları birleştirme
merged_df = pd.concat(dfs, ignore_index=True)

# Belirli başlıklar dışındaki sütunları silme
selected_columns = ["UrunAdi", "StokAdedi", "AlisFiyati", "SatisFiyati", "Resim", "AramaTerimleri", "MorhipoKodu", "VaryasyonMorhipoKodu", "HepsiBuradaKodu"]
filtered_df = merged_df[selected_columns]

# Sonuç DataFrame'i tek bir Excel dosyasına yazma
filtered_df.to_excel("sonuc_excel.xlsx", index=False)














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





















# sonuc_excel dosyasını oku
sonuc_excel_file = "sonuc_excel.xlsx"
sonuc_df = pd.read_excel(sonuc_excel_file)

# GMT ve SİTA dosyasını oku
gmt_sita_df = pd.read_excel("GMT ve SİTA.xlsx")

# İlk adım: 'UrunAdi' sütunuyla eşleşme ve GMT/SİTA Etopla değerlerini al
def match_and_update(row):
    urun_adi = row['UrunAdi']
    
    # GMT Ürün Adı sütununda arama yap ve 'GMT Etopla' değerini al
    gmt_row = gmt_sita_df[gmt_sita_df['GMT Ürün Adı'] == urun_adi]
    if not gmt_row.empty:
        row['GMT Etopla'] = gmt_row['GMT Etopla'].values[0]
    else:
        row['GMT Etopla'] = None

    # SİTA Ürün Adı sütununda arama yap ve 'SİTA Etopla' değerini al
    sista_row = gmt_sita_df[gmt_sita_df['SİTA Ürün Adı'] == urun_adi]
    if not sista_row.empty:
        row['SİTA Etopla'] = sista_row['SİTA Etopla'].values[0]
    else:
        row['SİTA Etopla'] = None

    return row

# İlk adımı uygulayın
sonuc_df = sonuc_df.apply(match_and_update, axis=1)

# İkinci adım: 'UrunAdi Duzenleme' sütununa göre arama, sadece GMT Etopla veya SİTA Etopla boş ya da sıfırsa
def match_and_update_with_code(row):
    urun_kodu = row['UrunAdi Duzenleme']
    
    # Eğer GMT Etopla boş ya da sıfırsa, 'GMT Ürün Kodu' ile arama yap
    if pd.isna(row['GMT Etopla']) or row['GMT Etopla'] == 0:
        gmt_code_row = gmt_sita_df[gmt_sita_df['GMT Ürün Kodu'] == urun_kodu]
        if not gmt_code_row.empty:
            gmt_etopla = gmt_code_row['GMT Etopla'].values[0]
            row['GMT Etopla'] = "GMT'de Var" if gmt_etopla > 0 else gmt_etopla
    
    # Eğer SİTA Etopla boş ya da sıfırsa, 'SİTA Ürün Kodu' ile arama yap
    if pd.isna(row['SİTA Etopla']) or row['SİTA Etopla'] == 0:
        sista_code_row = gmt_sita_df[gmt_sita_df['SİTA Ürün Kodu'] == urun_kodu]
        if not sista_code_row.empty:
            sita_etopla = sista_code_row['SİTA Etopla'].values[0]
            row['SİTA Etopla'] = "SİTA'da Var" if sita_etopla > 0 else sita_etopla

    return row

# İkinci adımı uygulayın
sonuc_df = sonuc_df.apply(match_and_update_with_code, axis=1)

# Güncellenmiş DataFrame'i yeni bir Excel dosyasına kaydedin
updated_excel_file = "sonuc_excel.xlsx"
sonuc_df.to_excel(updated_excel_file, index=False)











# "sonuc_excel.xlsx" Excel dosyasını oku
df_calisma_alani = pd.read_excel('sonuc_excel.xlsx')

# Aynı "UrunAdi" hücrelerinin "StokAdedi" sayılarını toplama
df_calisma_alani.loc[:, "StokAdedi"] = df_calisma_alani.groupby("UrunAdi")["StokAdedi"].transform("sum")

# "VaryasyonMorhipoKodu" sütununun adını değiştirme
df_calisma_alani = df_calisma_alani.rename(columns={"VaryasyonMorhipoKodu": "N11 & Zimmet"})

# Hesaplamalarda metinsel verileri sıfır olarak ele almak için sayısal değerlere dönüştürme
# Orijinal veri bozulmadan yalnızca matematiksel işlemler için geçici sütunlar kullanılıyor
gmt_numeric = pd.to_numeric(df_calisma_alani["GMT Etopla"], errors="coerce").fillna(0)
sita_numeric = pd.to_numeric(df_calisma_alani["SİTA Etopla"], errors="coerce").fillna(0)
stok_adedi_numeric = pd.to_numeric(df_calisma_alani["StokAdedi"], errors="coerce").fillna(0)
n11_zimmet_numeric = pd.to_numeric(df_calisma_alani["N11 & Zimmet"], errors="coerce").fillna(0)

# "Toplam Stok Adedi" sütunlarını oluşturma
df_calisma_alani["Toplam Stok Adedi Her Şey Dahil"] = stok_adedi_numeric + n11_zimmet_numeric + gmt_numeric + sita_numeric
df_calisma_alani["Toplam Stok Adedi Site ve Diğer Depolar"] = stok_adedi_numeric + n11_zimmet_numeric

# Eksik değerleri sıfır ile doldurma (diğer sütunlar için)
df_calisma_alani['StokAdedi'].fillna(0, inplace=True)
df_calisma_alani['N11 & Zimmet'].fillna(0, inplace=True)
df_calisma_alani['GMT Etopla'].fillna(0, inplace=True)
df_calisma_alani['SİTA Etopla'].fillna(0, inplace=True)

# Güncellenmiş DataFrame'i yeni bir Excel dosyasına kaydet
updated_excel_file = "sonuc_excel.xlsx"
df_calisma_alani.to_excel(updated_excel_file, index=False)








# "MorhipoKodu" sütununun adını değiştirme /Komplo orduların
df_calisma_alani = df_calisma_alani.rename(columns={"MorhipoKodu": "Günlük Satış Adedi"})
df_calisma_alani['Günlük Satış Adedi'].fillna(0, inplace=True)

# "Kaç Güne Biter" sütununu oluşturma ve "Toplam Stok Adedi" sütunundaki verileri "Günlük Satış Adedi" sütunundaki verilere bölme işlemi
df_calisma_alani["Kaç Güne Biter Her Şey Dahil"] = "Satış Adedi Yok"  # Varsayılan değer olarak "Satış Adedi Yok" atanır
df_calisma_alani["Kaç Güne Biter Site ve Diğer Depolar"] = "Satış Adedi Yok"  # Varsayılan değer olarak "Satış Adedi Yok" atanır


non_zero_mask = df_calisma_alani["Günlük Satış Adedi"] != 0
df_calisma_alani.loc[non_zero_mask, "Kaç Güne Biter Her Şey Dahil"] = round(df_calisma_alani["Toplam Stok Adedi Her Şey Dahil"] / df_calisma_alani["Günlük Satış Adedi"])


non_zero_mask = df_calisma_alani["Günlük Satış Adedi"] != 0
df_calisma_alani.loc[non_zero_mask, "Kaç Güne Biter Site ve Diğer Depolar"] = round(df_calisma_alani["Toplam Stok Adedi Site ve Diğer Depolar"] / df_calisma_alani["Günlük Satış Adedi"])












# "Görüntülenmenin Satışa Dönüş Oranı" sütunu
df_calisma_alani["Görüntülenmenin Satışa Dönüş Oranı"] = "0"  # Varsayılan değer olarak "Satış Yok" atanır
df_calisma_alani = df_calisma_alani.rename(columns={"HepsiBuradaKodu": "Görüntülenme Adedi"})
df_calisma_alani['Görüntülenme Adedi'].fillna(0, inplace=True)
non_zero_mask = df_calisma_alani["Görüntülenme Adedi"] != 0
df_calisma_alani.loc[non_zero_mask, "Görüntülenmenin Satışa Dönüş Oranı"] = round((df_calisma_alani["Günlük Satış Adedi"] / df_calisma_alani["Görüntülenme Adedi"]) * 100, 2)



# Değişiklikleri kaydetmek için dosyayı yeniden yaz
df_calisma_alani.to_excel("sonuc_excel.xlsx", index=False)





#region Satış Raporu Tarihini Düne Göre Ayarlama

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
    # ChromeDriver'ı en son sürümüyle otomatik olarak indirip kullan
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    login_url = "https://task.haydigiy.com/kullanici-giris/?ReturnUrl=%2Fadmin"
    driver.get(login_url)

    # Giriş bilgilerini doldurma
    email_input = driver.find_element("id", "EmailOrPhone")
    email_input.send_keys("mustafa_kod@haydigiy.com")

    password_input = driver.find_element("id", "Password")
    password_input.send_keys("123456")
    password_input.send_keys(Keys.RETURN)

    # Belirttiğiniz sayfaya yönlendirme
    desired_page_url = "https://task.haydigiy.com/admin/exportorder/edit/154/"
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

#endregion

#region Satış Raporunu İndirme

# Excel dosyasının indirileceği URL
url = "https://task.haydigiy.com/FaprikaOrderXls/GZPCKE/1/"
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

#endregion

#region Adet Sütununu Sayıya Çevirme

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

#endregion

#region ToplamFiyat Sütununu Sayıya Çevirme

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

#endregion

#region Adet ve ToplamFiyat Sütununa ETOPLA yapma

# Excel dosyasını tekrar okumak
df = pd.read_excel("Satış Raporu.xlsx")

# UrunAdi sütununa göre gruplandırma ve Adet ile ToplamFiyat sütunlarındaki verileri toplama
df_grouped = df.groupby('UrunAdi', as_index=False).agg({
    'Adet': 'sum',
    'ToplamFiyat': 'sum'
})

# Düzenlenmiş dosyayı aynı adla kaydetmek
df_grouped.to_excel("Satış Raporu.xlsx", index=False)

#endregion

#region Öne Çıkanlar Listesine Satış Adetleri Listesindeki Verileri Çektirme

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

#endregion














































# Excel dosyasını oku
df_calisma_alani = pd.read_excel("sonuc_excel.xlsx")

# "Resim" sütunundaki ".jpeg" ifadesinden sonrasını temizleme ve ".jpeg" ekleme
df_calisma_alani["Resim"] = df_calisma_alani["Resim"].str.replace(r"\.jpeg.*$", "", regex=True) + ".jpeg"

# Resim bağlantılarını bir listeye kaydet
links = df_calisma_alani["Resim"].tolist()

# "StokAdedi" sütununun adını değiştirme
df_calisma_alani = df_calisma_alani.rename(columns={"StokAdedi": "İnstagram Stok Adedi"})

# Sütun sıralamasını ayarlama
column_order = ["UrunAdi", "İnstagram Stok Adedi", "Toplam Stok Adedi Her Şey Dahil", "Toplam Stok Adedi Site ve Diğer Depolar", 
                "Günlük Satış Adedi", "Dünün Satış Adedi", "Görüntülenme Adedi", "Görüntülenmenin Satışa Dönüş Oranı", 
                "Kaç Güne Biter Her Şey Dahil", "Kaç Güne Biter Site ve Diğer Depolar", "AlisFiyati", "SatisFiyati", 
                "AramaTerimleri", "Resim", "GMT Etopla", "SİTA Etopla"]
df_calisma_alani = df_calisma_alani[column_order]

# Tekrarlanan satırları silme
df_calisma_alani = df_calisma_alani.drop_duplicates(subset=["UrunAdi"])

# NaN değerlerini 0 ile değiştirme
df_calisma_alani = df_calisma_alani.fillna(0)

# inf değerlerini 0 ile değiştirme
df_calisma_alani.replace([float('inf'), float('-inf')], 0, inplace=True)

# "AramaTerimleri" sütunundaki tarihleri temizle
date_pattern = r'(\d{1,2}\.\d{1,2}\.\d{4})'
df_calisma_alani['AramaTerimleri'] = df_calisma_alani['AramaTerimleri'].apply(lambda x: re.search(date_pattern, str(x)).group(1) if re.search(date_pattern, str(x)) else None)

# "Resim" sütununu DataFrame'den kaldır
df_calisma_alani.drop(columns=["Resim"], inplace=True)








# Güncellenmiş DataFrame'i aynı Excel dosyasına yaz
with pd.ExcelWriter('sonuc_excel.xlsx', engine='xlsxwriter') as writer:
    df_calisma_alani.to_excel(writer, index=False, sheet_name='Sheet1')

    # ExcelWriter objesinden workbook ve worksheet'e eriş
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # İlk sütun (UrunAdi) için uygun genişlik ayarlama
    max_col_width = max(df_calisma_alani["UrunAdi"].astype(str).apply(len).max(), len("UrunAdi")) + 2
    worksheet.set_column(0, 0, max_col_width)

    # Belirli sütunlar için genişliği 10 piksel olarak ayarlama
    narrow_columns = ["AlisFiyati", "SatisFiyati", "GMT Etopla", "SİTA Etopla"]
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

    # Satırları ve sütunları dolaşarak biçimlendirme
    for row_num, row in enumerate(df_calisma_alani.itertuples(), start=1):
        for col_num, value in enumerate(row[1:]):  # row[0] index olduğu için atlanıyor
            col_name = df_calisma_alani.columns[col_num]
            if col_num == 0:  # UrunAdi sütunu
                link = links[row_num - 1]  # Resim sütunundaki bağlantıyı UrunAdi'ye ekliyoruz
                if row_num % 2 == 1:
                    worksheet.write_url(row_num, col_num, link, string=value, cell_format=shaded_left_align_format)
                else:
                    worksheet.write_url(row_num, col_num, link, string=value, cell_format=left_align_format)
            elif col_name in ["AlisFiyati", "SatisFiyati"]:
                # Para birimi formatı uygulama, alternatif satır renklendirme ile ve orta hizalı
                if row_num % 2 == 1:
                    worksheet.write(row_num, col_num, value, shaded_currency_format)
                else:
                    worksheet.write(row_num, col_num, value, currency_format)
            elif col_name in ["GMT Etopla", "SİTA Etopla"] and "Var" in str(value):
                worksheet.write(row_num, col_num, value, var_format)
            else:
                if row_num % 2 == 1:
                    worksheet.write(row_num, col_num, value, shaded_format)
                else:
                    worksheet.write(row_num, col_num, value, center_format)

    # Başlıklara filtre ekleme
    worksheet.autofilter(0, 0, 0, len(df_calisma_alani.columns) - 1)

    # Sayfanın yakınlaştırma oranını %90 olarak ayarla
    worksheet.set_zoom(90)

# Dosyanın adını değiştirme
excel_file_name = "sonuc_excel.xlsx"
new_excel_file_name = "Nirvana.xlsx"
os.rename(excel_file_name, new_excel_file_name)

# Eski dosyaları silme
dosyalar = ["GMT ve SİTA.xlsx", "Satış Raporu.xlsx"]

for dosya in dosyalar:
    if os.path.exists(dosya):
        os.remove(dosya)
