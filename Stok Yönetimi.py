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


print("kalsaydın olmaz mıydı")

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
            print("Dosyalar zaten silinmiş veya bulunamadı.")
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




# Google Sheet URL
google_sheet_url = "https://docs.google.com/spreadsheets/d/1aA5LhkQYgtwHLcKRV1mKl9Lb6VeOgUNIC9zy2kRagrs/gviz/tq?tqx=out:csv"

# Google Sheet'ten veriyi al ve Excel dosyasına kaydet
try:
    google_df = pd.read_csv(google_sheet_url)
    google_excel_file = "GMT ve SİTA.xlsx"
    google_df.to_excel(google_excel_file, index=False)
except requests.exceptions.RequestException as e:
    print(f"Request failed: {e}")

# sonuc_excel dosyasını oku
sonuc_excel_file = "sonuc_excel.xlsx"
sonuc_df = pd.read_excel(sonuc_excel_file)

# GMT ve SİTA dosyasını oku
gmt_sita_df = pd.read_excel("GMT ve SİTA.xlsx")

# 'UrunAdi' sütunundaki her değeri işle
def match_and_update(row):
    urun_adi = row['UrunAdi']
    
    # GMT ve SİTA dosyasında 'GMT Ürün Adı' sütununda arama yap ve 'GMT Etopla' değerini al
    gmt_row = gmt_sita_df[gmt_sita_df['GMT Ürün Adı'] == urun_adi]
    if not gmt_row.empty:
        row['GMT Etopla'] = gmt_row['GMT Etopla'].values[0]
    else:
        row['GMT Etopla'] = None

    # GMT ve SİTA dosyasında 'İSTA Ürün Adı' sütununda arama yap ve 'SİTA Etopla' değerini al
    sista_row = gmt_sita_df[gmt_sita_df['SİTA Ürün Adı'] == urun_adi]
    if not sista_row.empty:
        row['SİTA Etopla'] = sista_row['SİTA Etopla'].values[0]
    else:
        row['SİTA Etopla'] = None

    return row

# 'UrunAdi' sütununu kullanarak işlemi gerçekleştirin
sonuc_df = sonuc_df.apply(match_and_update, axis=1)

# Güncellenmiş DataFrame'i yeni bir Excel dosyasına kaydedin
updated_excel_file = "sonuc_excel.xlsx"
sonuc_df.to_excel(updated_excel_file, index=False)











# "sonuc_excel.xlsx" Excel dosyasını oku
df_calisma_alani = pd.read_excel('sonuc_excel.xlsx')

# Aynı "UrunAdi" hücrelerinin "StokAdedi" sayılarını toplama
df_calisma_alani.loc[:, "StokAdedi"] = df_calisma_alani.groupby("UrunAdi")["StokAdedi"].transform("sum")

# "VaryasyonMorhipoKodu" sütununun adını değiştirme
df_calisma_alani = df_calisma_alani.rename(columns={"VaryasyonMorhipoKodu": "N11 & Zimmet"})

# Veri tiplerini uyumlu hale getirme
df_calisma_alani["StokAdedi"] = pd.to_numeric(df_calisma_alani["StokAdedi"], errors="coerce")
df_calisma_alani["N11 & Zimmet"] = pd.to_numeric(df_calisma_alani["N11 & Zimmet"], errors="coerce")
df_calisma_alani["GMT Etopla"] = pd.to_numeric(df_calisma_alani["GMT Etopla"], errors="coerce")
df_calisma_alani["SİTA Etopla"] = pd.to_numeric(df_calisma_alani["SİTA Etopla"], errors="coerce")

# Eksik değerleri sıfır ile doldurma
df_calisma_alani['StokAdedi'].fillna(0, inplace=True)
df_calisma_alani['N11 & Zimmet'].fillna(0, inplace=True)
df_calisma_alani['GMT Etopla'].fillna(0, inplace=True)
df_calisma_alani['SİTA Etopla'].fillna(0, inplace=True)

# "Toplam Stok Adedi" sütununu oluştur
df_calisma_alani["Toplam Stok Adedi Her Şey Dahil"] = df_calisma_alani["StokAdedi"] + df_calisma_alani["N11 & Zimmet"] + df_calisma_alani["GMT Etopla"] + df_calisma_alani["SİTA Etopla"]

# "Toplam Stok Adedi" sütununu oluştur
df_calisma_alani["Toplam Stok Adedi Site ve Diğer Depolar"] = df_calisma_alani["StokAdedi"] + df_calisma_alani["N11 & Zimmet"]








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

# "Resim" sütunundaki ".jpeg" ifadesinden sonrasını temizleme
df_calisma_alani["Resim"] = df_calisma_alani["Resim"].str.replace(r"\.jpeg.*$", "", regex=True)

# Kalan verilere ".jpeg" eklenmesi
df_calisma_alani["Resim"] = df_calisma_alani["Resim"] + ".jpeg"

# "StokAdedi" sütununun adını değiştirme
df_calisma_alani = df_calisma_alani.rename(columns={"StokAdedi": "İnstagram Stok Adedi"})

# Sütun sıralamasını ayarlama
column_order = ["UrunAdi", "İnstagram Stok Adedi", "Toplam Stok Adedi Her Şey Dahil", "Toplam Stok Adedi Site ve Diğer Depolar", "Günlük Satış Adedi", "Görüntülenme Adedi", "Görüntülenmenin Satışa Dönüş Oranı", "Kaç Güne Biter Her Şey Dahil", "Kaç Güne Biter Site ve Diğer Depolar", "AlisFiyati", "SatisFiyati", "AramaTerimleri", "Resim", "GMT Etopla", "SİTA Etopla"]
df_calisma_alani = df_calisma_alani[column_order]

# Tekrarlanan satırları silme
df_calisma_alani = df_calisma_alani.drop_duplicates(subset=["UrunAdi"])

# NaN değerlerini 0 ile değiştirme
df_calisma_alani = df_calisma_alani.fillna(0)

# inf değerlerini 0 ile değiştirme
df_calisma_alani.replace([float('inf'), float('-inf')], 0, inplace=True)

# Sonuç DataFrame'ini tekrar "sonuc_excel.xlsx" adlı bir Excel dosyasına yazma
df_calisma_alani.to_excel("sonuc_excel.xlsx", index=False)

# Tarihleri çıkarmak için regex deseni
date_pattern = r'(\d{1,2}\.\d{1,2}\.\d{4})'

# "AramaTerimleri" sütunundaki tarihleri temizle
df_calisma_alani['AramaTerimleri'] = df_calisma_alani['AramaTerimleri'].apply(lambda x: re.search(date_pattern, str(x)).group(1) if re.search(date_pattern, str(x)) else None)

# Güncellenmiş DataFrame'i aynı Excel dosyasının üzerine yaz
with pd.ExcelWriter('sonuc_excel.xlsx', engine='xlsxwriter') as writer:
    df_calisma_alani.to_excel(writer, index=False, sheet_name='Sheet1')

    # ExcelWriter objesinden workbook ve worksheet'e eriş
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    # DataFrame sütun genişliklerini al
    column_widths = [max(df_calisma_alani[col].astype(str).apply(len).max(), len(col)) + 2 for col in df_calisma_alani.columns]

    # Sütun genişliklerini Excel worksheet'e ayarla
    for i, width in enumerate(column_widths):
        worksheet.set_column(i, i, width)

    # Tabloyu ortala
    center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
    header_format = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'align': 'center', 'valign': 'vcenter'})

    for col_num, value in enumerate(df_calisma_alani.columns.values):
        worksheet.write(0, col_num, value, header_format)

    for i, col in enumerate(df_calisma_alani.columns):
        for j, value in enumerate(df_calisma_alani[col]):
            if col == 'Resim':
                worksheet.write_url(j + 1, i, value, string='Tıkla', cell_format=center_format)
            else:
                worksheet.write(j + 1, i, value, center_format)

    # "Resim" sütununun genişliğini 20 piksel olarak ayarla
    worksheet.set_column('M:M', 20)

# Dosyanın adını değiştirme
excel_file_name = "sonuc_excel.xlsx"
new_excel_file_name = "Satış Raporu.xlsx"
os.rename(excel_file_name, new_excel_file_name)

# Eski dosyaları silme
dosyalar = ["GMT ve SİTA.xlsx"]

for dosya in dosyalar:
    if os.path.exists(dosya):
        os.remove(dosya)

