#region // Kütüphaneler

import requests
from bs4 import BeautifulSoup
import pandas as pd
from io import BytesIO
import re
from colorama import init, Fore, Style
from datetime import datetime, timedelta
import os
import openpyxl
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill
import xlsxwriter
import gc
from supabase import create_client, Client
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
import http.client
import json
import datetime
import warnings
import copy
import colorama
from colorama import Fore, Style
init(autoreset=True)
warnings.filterwarnings("ignore")
colorama.init(autoreset=True)

#endregion

#region // Entegrasyondan Önce mi Sonra mı Kontrolü ve Satış Raporu Tarihini Düne Göre Ayarlama










def list_detail_with_http_client():
    # 1) Giriş (login) isteği
    conn = http.client.HTTPSConnection("siparis.haydigiy.com")

    login_payload = {
        "apiKey": "MypGcaEInEOTzuYQydgDHQ",
        "secretKey": "jRqliBLDPke76YhL_WL5qg",
        "emailOrPhone": "mustafa_kod@haydigiy.com",
        "password": "123456"
    }
    login_headers = {
        "Content-Type": "application/json",
        "Accept": "application/json"
    }

    conn.request(
        "POST",
        "/api/customer/login",
        body=json.dumps(login_payload),
        headers=login_headers
    )

    res = conn.getresponse()
    data = res.read().decode("utf-8")

    if res.status != 200:
        print("Giriş başarısız:", data)
        return

    login_data = json.loads(data)
    token = login_data.get("data", {}).get("token")
    if not token:
        print("Token alınamadı. Dönen veri:", login_data)
        return

    # 2) list-detail isteği
    conn = http.client.HTTPSConnection("siparis.haydigiy.com")

    list_detail_payload = {
        "searchTerm": None,
        "inCategoryIds": [],
        "includeInSubCategories": True,
        "notInCategoryIds": [],
        "includeNotInSubCategories": True,
        "inManufacturerIds": [],
        "notInManufacturerIds": [],
        "inVendorIds": [],
        "notInVendorIds": [],
        "inProductTagIds": [],
        "notInProductTagIds": [],
        "published": None,
        "minStock": None,
        "maxStock": None,
        "minPrice": None,
        "maxPrice": None,
        "minProductCost": None,
        "maxProductCost": None,
        "pageIndex": 1,
        "pageSize": 1
    }

    list_detail_headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Accept": "application/json"
    }

    conn.request(
        "POST",
        "/adminapi/product/list-detail",
        body=json.dumps(list_detail_payload),
        headers=list_detail_headers
    )

    res = conn.getresponse()
    data = res.read().decode("utf-8")

    if res.status != 200:
        print("List-detail isteği başarısız:", data)
        return

    # 3) Gelen JSON'u parse edelim
    response_data = json.loads(data)

    items = response_data.get("data", [])
    if not items:
        print("Herhangi bir ürün kaydı bulunamadı.")
        return

    first_item = items[0]
    created_on_str = first_item.get("createdOn")
    if not created_on_str:
        print("İlk üründe 'createdOn' alanı yok.")
        return

    # 4) createdOn değerini ISO formatından datetime'a çevirelim
    try:
        created_on_dt = datetime.datetime.fromisoformat(created_on_str.replace("Z", ""))
        today = datetime.datetime.now().date()
        if created_on_dt.date() == today:
            print("\033[92mEntegrasyondan Sonraki Listeyi Çekiyorsunuz !\033[0m")
        else:
            print("\033[91mDikkat Entegrasyondan Önceki Listeyi Çekiyorsunuz !\033[0m")

    except ValueError:
        print("createdOn alanı beklenmeyen bir formatta:", created_on_str)

if __name__ == "__main__":
    list_detail_with_http_client()



















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
desired_page_url = "https://www.siparis.haydigiy.com/admin/exportorder/edit/154/"

try:
    # Giriş sayfasına git
    driver.get(login_url)
    time.sleep(2)  # Sayfanın yüklenmesini bekleyin

    # Giriş bilgilerini doldur
    driver.find_element(By.NAME, "EmailOrPhone").send_keys(username)
    driver.find_element(By.NAME, "Password").send_keys(password)
    driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
    time.sleep(3)  # Giriş sonrası bekleme süresi

    # Belirtilen sayfaya git
    driver.get(desired_page_url)
    time.sleep(2)

    # Dünün tarihini (gün ve ay için başında sıfır olmadan) alalım
    yesterday = datetime.now() - timedelta(days=1)
    formatted_date_no_leading = f"{yesterday.day}.{yesterday.month}.{yesterday.year}"

    # EndDate alanını bulma ve tarih girişini yapma
    end_date_input = driver.find_element(By.ID, "EndDate")
    end_date_input.clear()
    end_date_input.send_keys(formatted_date_no_leading)

    # StartDate alanını bulma ve tarih girişini yapma
    start_date_input = driver.find_element(By.ID, "StartDate")
    start_date_input.clear()
    start_date_input.send_keys(formatted_date_no_leading)

    # Kaydet butonunu bulma ve tıklama
    save_button = driver.find_element(By.CSS_SELECTOR, 'button.btn.btn-primary[name="save"]')
    save_button.click()

    print(Fore.GREEN + "Tarih ayarlama işlemi başarılı!" + Style.RESET_ALL)

except Exception as e:
    print(Fore.RED + f"Hata oluştu: {e}" + Style.RESET_ALL)

finally:
    # Tarayıcıyı kapat
    driver.quit()
























#endregion

#region // GMT ve SİTA Verilerini Çekme

SUPABASE_URL = "https://zmvsatlvobhdaxxgtoap.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InptdnNhdGx2b2JoZGF4eGd0b2FwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDAxNzIxMzksImV4cCI6MjA1NTc0ODEzOX0.lJLudSfixMbEOkJmfv22MsRLofP7ZjFkbGj26xF3dts"

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

all_data = []
start = 0
page_size = 1000

while True:
    end = start + page_size - 1
    response = (
        supabase.table("urunyonetimi")
        .select("urunkodu, renk, acilmamisadet, gmtsitalabel")
        .in_("gmtsitalabel", ["GMT", "SİTA", "Yarım GMT"])
        .gt("acilmamisadet", 0)
        .range(start, end)
        .execute()
    )
    data = response.data
    if not data:
        break
    all_data.extend(data)
    start += page_size

df = pd.DataFrame(all_data)
df["renk"] = df["renk"].apply(lambda x: x.capitalize() if isinstance(x, str) else x)

df_gmt = df[df["gmtsitalabel"].isin(["GMT", "Yarım GMT"])]
df_gmt_grouped = df_gmt.groupby(["urunkodu", "renk"], as_index=False)["acilmamisadet"].sum()

df_gmt_final = pd.DataFrame()
df_gmt_final["GMT Ürün Kodu"] = df_gmt_grouped["urunkodu"]
df_gmt_final["GMT Ürün Adı"] = df_gmt_grouped["urunkodu"].astype(str) + " - " + df_gmt_grouped["renk"]
df_gmt_final["GMT Stok Adedi"] = df_gmt_grouped["acilmamisadet"]
df_gmt_final["GMT Ürün Kodu"] = pd.to_numeric(df_gmt_final["GMT Ürün Kodu"], errors="coerce").astype("Int64")

df_sita = df[df["gmtsitalabel"] == "SİTA"]
df_sita_grouped = df_sita.groupby(["urunkodu", "renk"], as_index=False)["acilmamisadet"].sum()

df_sita_final = pd.DataFrame()
df_sita_final["SİTA Ürün Kodu"] = df_sita_grouped["urunkodu"]
df_sita_final["SİTA Ürün Adı"] = df_sita_grouped["urunkodu"].astype(str) + " - " + df_sita_grouped["renk"]
df_sita_final["SİTA Stok Adedi"] = df_sita_grouped["acilmamisadet"]
df_sita_final["SİTA Ürün Kodu"] = pd.to_numeric(df_sita_final["SİTA Ürün Kodu"], errors="coerce").astype("Int64")

with pd.ExcelWriter("GMT ve SİTA.xlsx") as writer:
    df_gmt_final.to_excel(writer, sheet_name="Sheet1", index=False)
    df_sita_final.to_excel(writer, sheet_name="Sheet2", index=False)

#endregion

#region // GMT ve SİTA Verilerini Ana Tabloya Çektirme (Etopla Yapma)

one_cikanlar_df = pd.DataFrame()  # Burada ileride kullanacağımız boş bir DF tanımı (varsa hataya düşmemek için)

#endregion

#region // Ürün Listesi İndirme

SUPABASE_URL = "https://zmvsatlvobhdaxxgtoap.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InptdnNhdGx2b2JoZGF4eGd0b2FwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDAxNzIxMzksImV4cCI6MjA1NTc0ODEzOX0.lJLudSfixMbEOkJmfv22MsRLofP7ZjFkbGj26xF3dts"

# Supabase istemcisini oluştur
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# Storage üzerinden "tum_urun_listesi.xlsx" dosyasını indir
# BUCKET_ADINIZ kısmını ve dosya yolunu kendi projenize göre düzenleyin.
response = supabase.storage.from_("tum_urun_listesi").download("tum_urun_listesi.xlsx")

# Okunacak veriyi BytesIO aracılığıyla pandas'a aktar
temp_df = pd.read_excel(BytesIO(response))

# "UrunAdi" sütununda "-" karakteri içeren satırları filtreleyelim
filtered_rows = temp_df[
    temp_df["UrunAdi"].astype(str).str.contains(re.escape("-"), case=False, na=False)
]

# Filtrelenmiş DataFrame boş değilse belirli sütunları seçip çıktı alıyoruz
if not filtered_rows.empty:
    selected_columns = [
        "UrunAdi", 
        "StokAdedi", 
        "AlisFiyati", 
        "SatisFiyati", 
        "Kategori",
        "Resim", 
        "AramaTerimleri", 
        "MorhipoKodu", 
        "VaryasyonMorhipoKodu",
        "HepsiBuradaKodu", 
        "Marka", 
        "N11Kodu", 
        "VaryasyonGittiGidiyorKodu",
        "TrendyolKodu",
        "VaryasyonTrendyolKodu"
    ]
    final_df = filtered_rows[selected_columns]
    final_df.to_excel("sonuc_excel.xlsx", index=False)
else:
    # Hiç kayıt yoksa boş bir Excel oluştur
    pd.DataFrame().to_excel("sonuc_excel.xlsx", index=False)

#endregion

#region // Ürünlerin Kategorilerini Belirleme ve Tesettür Ayarlaması

df = pd.read_excel("sonuc_excel.xlsx")
df['Kategori'] = df['Kategori'].fillna("")

def extract_category(text):
    if not isinstance(text, str):
        return None
    match = re.search(r'>\s*([^;]+)', text)
    if match:
        return match.group(1).strip()
    elif "TESETTÜR" in text:
        return "TESETTÜR"
    return None

df['Kategori'] = df['Kategori'].apply(extract_category)
df.to_excel("sonuc_excel.xlsx", index=False)

#endregion

#region // UrunAdi Duzenleme Sütununu Oluşturma ve Sadece Ürün Kodlarını Bırakma

sonuc_df = pd.read_excel("sonuc_excel.xlsx")

def extract_product_code(urun_adi):
    match = re.search(r' - (\d+)\.', urun_adi)
    return match.group(1) if match else None

def extract_color(urun_adi):
    parts = re.split(r' - ', urun_adi)
    if len(parts) > 0:
        before_part = parts[0].strip()
        words = before_part.split()
        if words:
            return words[-1]
    return None

sonuc_df['UrunAdi Duzenleme'] = sonuc_df['UrunAdi'].apply(extract_product_code)
sonuc_df['UrunAdi Duzenleme'] = sonuc_df['UrunAdi Duzenleme'].astype(str)

sonuc_df['UrunAdi ve Renk'] = sonuc_df.apply(
    lambda row: row['UrunAdi Duzenleme'] + " - " + extract_color(row['UrunAdi']),
    axis=1
)

sonuc_df.to_excel("sonuc_excel.xlsx", index=False)

#endregion

#region // GMT ve SİTA Verilerini Ana Tabloya Çektirme (Etopla Yapma)

df_calisma_alani = pd.read_excel("sonuc_excel.xlsx")

gmt_df = pd.read_excel("GMT ve SİTA.xlsx", sheet_name="Sheet1")
sita_df = pd.read_excel("GMT ve SİTA.xlsx", sheet_name="Sheet2")

used_gmt_indices_step1 = []
used_sita_indices_step1 = []

for idx, row in df_calisma_alani.iterrows():
    urun_adi = row['UrunAdi ve Renk']
    
    matching_gmt = gmt_df[gmt_df['GMT Ürün Adı'] == urun_adi]
    if not matching_gmt.empty:
        matched_index = matching_gmt.index[0]
        df_calisma_alani.at[idx, 'GMT Stok Adedi'] = matching_gmt.iloc[0]['GMT Stok Adedi']
        used_gmt_indices_step1.append(matched_index)
    else:
        df_calisma_alani.at[idx, 'GMT Stok Adedi'] = None

    matching_sita = sita_df[sita_df['SİTA Ürün Adı'] == urun_adi]
    if not matching_sita.empty:
        matched_index = matching_sita.index[0]
        df_calisma_alani.at[idx, 'SİTA Stok Adedi'] = matching_sita.iloc[0]['SİTA Stok Adedi']
        used_sita_indices_step1.append(matched_index)
    else:
        df_calisma_alani.at[idx, 'SİTA Stok Adedi'] = None

gmt_df = gmt_df.drop(used_gmt_indices_step1).reset_index(drop=True)
sita_df = sita_df.drop(used_sita_indices_step1).reset_index(drop=True)

used_gmt_indices_step2 = []
used_sita_indices_step2 = []

for idx, row in df_calisma_alani.iterrows():
    urun_kodu = row['UrunAdi Duzenleme']
    
    if pd.isna(row.get('GMT Stok Adedi')) or row.get('GMT Stok Adedi') == 0:
        matching_gmt_code = gmt_df[gmt_df['GMT Ürün Kodu'] == urun_kodu]
        if not matching_gmt_code.empty:
            matched_index = matching_gmt_code.index[0]
            gmt_stok = matching_gmt_code.iloc[0]['GMT Stok Adedi']
            df_calisma_alani.at[idx, 'GMT Stok Adedi'] = "GMT'de Var" if gmt_stok > 0 else gmt_stok
            used_gmt_indices_step2.append(matched_index)

    if pd.isna(row.get('SİTA Stok Adedi')) or row.get('SİTA Stok Adedi') == 0:
        matching_sita_code = sita_df[sita_df['SİTA Ürün Kodu'] == urun_kodu]
        if not matching_sita_code.empty:
            matched_index = matching_sita_code.index[0]
            sita_stok = matching_sita_code.iloc[0]['SİTA Stok Adedi']
            df_calisma_alani.at[idx, 'SİTA Stok Adedi'] = "SİTA'da Var" if sita_stok > 0 else sita_stok
            used_sita_indices_step2.append(matched_index)

gmt_df = gmt_df.drop(used_gmt_indices_step2).reset_index(drop=True)
sita_df = sita_df.drop(used_sita_indices_step2).reset_index(drop=True)

df_calisma_alani.to_excel("sonuc_excel.xlsx", index=False)

with pd.ExcelWriter("GMT ve SİTA.xlsx", engine='openpyxl') as writer:
    gmt_df.to_excel(writer, sheet_name='Sheet1', index=False)
    sita_df.to_excel(writer, sheet_name='Sheet2', index=False)

#endregion

#region // Stok Adedi Sütunu İçin Etopla Yapma - Stok Adedi Her Şey Dahil ve Stok Adedi Site ve Vega Sütunlarını Oluşturma - Bazı Sütunların Adını Değiştirme

df_calisma_alani = pd.read_excel("sonuc_excel.xlsx")
df_calisma_alani.loc[:, "StokAdedi"] = df_calisma_alani.groupby("UrunAdi")["StokAdedi"].transform("sum")

df_calisma_alani = df_calisma_alani.rename(columns={"VaryasyonMorhipoKodu": "N11 & Zimmet"})

gmt_numeric = pd.to_numeric(df_calisma_alani["GMT Stok Adedi"], errors="coerce").fillna(0)
sita_numeric = pd.to_numeric(df_calisma_alani["SİTA Stok Adedi"], errors="coerce").fillna(0)
stok_adedi_numeric = pd.to_numeric(df_calisma_alani["StokAdedi"], errors="coerce").fillna(0)
n11_zimmet_numeric = pd.to_numeric(df_calisma_alani["N11 & Zimmet"], errors="coerce").fillna(0)

df_calisma_alani["Stok Adedi Her Şey Dahil"] = stok_adedi_numeric + n11_zimmet_numeric + gmt_numeric + sita_numeric
df_calisma_alani["Stok Adedi Site ve Vega"] = stok_adedi_numeric + n11_zimmet_numeric

df_calisma_alani['StokAdedi'].fillna(0, inplace=True)
df_calisma_alani['N11 & Zimmet'].fillna(0, inplace=True)
df_calisma_alani['GMT Stok Adedi'].fillna(0, inplace=True)
df_calisma_alani['SİTA Stok Adedi'].fillna(0, inplace=True)

df_calisma_alani.to_excel("sonuc_excel.xlsx", index=False)

#endregion

#region // MorhipoKodu Sütununun Adını Değiştirme ve Kaç Güne Biter Kısımlarını Hesaplama

df_calisma_alani = pd.read_excel("sonuc_excel.xlsx")
df_calisma_alani = df_calisma_alani.rename(columns={"MorhipoKodu": "Günlük Ortalama Satış Adedi"})
df_calisma_alani['Günlük Ortalama Satış Adedi'].fillna(0, inplace=True)

df_calisma_alani["Kaç Güne Biter Her Şey Dahil"] = "Satış Adedi Yok"
df_calisma_alani["Kaç Güne Biter Site ve Vega"] = "Satış Adedi Yok"

non_zero_mask = df_calisma_alani["Günlük Ortalama Satış Adedi"] != 0
df_calisma_alani.loc[non_zero_mask, "Kaç Güne Biter Her Şey Dahil"] = round(
    df_calisma_alani["Stok Adedi Her Şey Dahil"] / df_calisma_alani["Günlük Ortalama Satış Adedi"]
)

non_zero_mask = df_calisma_alani["Günlük Ortalama Satış Adedi"] != 0
df_calisma_alani.loc[non_zero_mask, "Kaç Güne Biter Site ve Vega"] = round(
    df_calisma_alani["Stok Adedi Site ve Vega"] / df_calisma_alani["Günlük Ortalama Satış Adedi"]
)

df_calisma_alani.to_excel("sonuc_excel.xlsx", index=False)

#endregion

#region // Görüntülenmenin Satışa Dönüş Oranını Hesaplama

df_calisma_alani = pd.read_excel("sonuc_excel.xlsx")
df_calisma_alani["Görüntülenmenin Satışa Dönüş Oranı"] = "0"
df_calisma_alani = df_calisma_alani.rename(columns={"HepsiBuradaKodu": "Ortalama Görüntülenme Adedi"})
df_calisma_alani['Ortalama Görüntülenme Adedi'].fillna(0, inplace=True)

non_zero_mask = df_calisma_alani["Ortalama Görüntülenme Adedi"] != 0
df_calisma_alani.loc[non_zero_mask, "Görüntülenmenin Satışa Dönüş Oranı"] = round(
    (df_calisma_alani["Günlük Ortalama Satış Adedi"] / df_calisma_alani["Ortalama Görüntülenme Adedi"]) * 100, 2
)

df_calisma_alani.to_excel("sonuc_excel.xlsx", index=False)

#endregion

#region // Satış Raporunu İndirme

url = "https://www.siparis.haydigiy.com/FaprikaOrderXls/GZPCKE/1/"
filename = "Satış Raporu.xlsx"

def is_file_downloaded_today(file_path):
    if os.path.exists(file_path):
        file_mod_time = os.path.getmtime(file_path)
        mod_date = datetime.fromtimestamp(file_mod_time).date()
        return mod_date == datetime.today().date()
    return False

if not is_file_downloaded_today(filename):
    if os.path.exists(filename):
        os.remove(filename)
    resp = requests.get(url)
    with open(filename, "wb") as f:
        f.write(resp.content)

df = pd.read_excel(filename)
columns_to_keep = ["UrunAdi", "Adet", "ToplamFiyat"]
df = df[columns_to_keep]
df.to_excel(filename, index=False)

#endregion

#region // Adet Sütununu Sayıya Çevirme

def clean_adet(data):
    data['Adet'] = data['Adet'].astype(str).apply(lambda x: x.split(',')[0])

def convert_adet_to_numeric(data):
    data['Adet'] = pd.to_numeric(data['Adet'], errors='coerce')

if __name__ == "__main__":
    file_path = "Satış Raporu.xlsx"
    combined_data = pd.read_excel(file_path, engine="openpyxl")
    clean_adet(combined_data)
    combined_data.to_excel(file_path, index=False, engine='openpyxl')

    combined_data = pd.read_excel(file_path, engine="openpyxl")
    convert_adet_to_numeric(combined_data)
    combined_data.to_excel(file_path, index=False, engine='openpyxl')

#endregion

#region // ToplamFiyat Sütununu Sayıya Çevirme

def clean_toplamfiyat(data):
    data['ToplamFiyat'] = data['ToplamFiyat'].astype(str).apply(lambda x: x.split(',')[0])

def convert_toplamfiyat_to_numeric(data):
    data['ToplamFiyat'] = pd.to_numeric(data['ToplamFiyat'], errors='coerce')

if __name__ == "__main__":
    file_path = "Satış Raporu.xlsx"
    combined_data = pd.read_excel(file_path, engine="openpyxl")
    clean_toplamfiyat(combined_data)
    combined_data.to_excel(file_path, index=False, engine='openpyxl')

    combined_data = pd.read_excel(file_path, engine="openpyxl")
    convert_toplamfiyat_to_numeric(combined_data)
    combined_data.to_excel(file_path, index=False, engine='openpyxl')

#endregion

#region // Adet ve ToplamFiyat Sütununa ETOPLA yapma

df = pd.read_excel("Satış Raporu.xlsx")
df_grouped = df.groupby('UrunAdi', as_index=False).agg({
    'Adet': 'sum',
    'ToplamFiyat': 'sum'
})
df_grouped.to_excel("Satış Raporu.xlsx", index=False)

#endregion

#region // Ana Listeye Veriyi Çektirme

satis_raporu_df = pd.read_excel("Satış Raporu.xlsx")
one_cikanlar_df = pd.read_excel("sonuc_excel.xlsx")

merged_df = one_cikanlar_df.merge(
    satis_raporu_df[['UrunAdi', 'Adet']],
    on='UrunAdi',
    how='left'
)

merged_df.rename(columns={'Adet': 'Dünün Satış Adedi'}, inplace=True)
merged_df.to_excel("sonuc_excel.xlsx", index=False)

#endregion

#region // Sütunların Sırasını Değiştirme - Bazı Sütunların Adını Değiştirme

df_calisma_alani = pd.read_excel("sonuc_excel.xlsx")

# İsim değişiklikleri
df_calisma_alani = df_calisma_alani.rename(columns={"StokAdedi": "İnstagram Stok Adedi"})
df_calisma_alani = df_calisma_alani.rename(columns={"VaryasyonGittiGidiyorKodu": "Net Satış Tarihi ve Adedi"})
df_calisma_alani = df_calisma_alani.rename(columns={"TrendyolKodu": "Son Transfer Tarihi"})
df_calisma_alani = df_calisma_alani.rename(columns={"VaryasyonTrendyolKodu": "Son İndirim Tarihi"})

# Sütun sırası
column_order = [
    "UrunAdi",
    "İnstagram Stok Adedi",
    "Stok Adedi Her Şey Dahil",
    "Stok Adedi Site ve Vega",
    "Günlük Ortalama Satış Adedi",
    "Dünün Satış Adedi",
    "Ortalama Görüntülenme Adedi",
    "Görüntülenmenin Satışa Dönüş Oranı",
    "Kaç Güne Biter Her Şey Dahil",
    "Kaç Güne Biter Site ve Vega",
    "AlisFiyati",
    "SatisFiyati",
    "AramaTerimleri",
    "Resim",
    "Kategori",
    "GMT Stok Adedi",
    "SİTA Stok Adedi",
    "Marka",
    "N11Kodu",
    "Net Satış Tarihi ve Adedi",
    "Son Transfer Tarihi",
    "Son İndirim Tarihi"
]
df_calisma_alani = df_calisma_alani[column_order]

# "Son İndirim Tarihi" ve "Son Transfer Tarihi" kolonlarındaki verilerin boşluktan sonraki kısımlarını silelim
for col in ["Son İndirim Tarihi", "Son Transfer Tarihi"]:
    df_calisma_alani[col] = (
        df_calisma_alani[col]
        .astype(str)  # metne çevir
        .apply(lambda x: x.split(' ')[0] if x and x.lower() != 'nan' else '')
    )

df_calisma_alani.to_excel("sonuc_excel.xlsx", index=False)

#endregion

#region // Yenilenen Değerleri Kaldırma

df_calisma_alani = df_calisma_alani.drop_duplicates(subset=["UrunAdi"])

#endregion

#region // Resim Sütunu İçin .jpeg'den Sonrasını Kaldırma ve Devamına .jpeg Ekleme

df_calisma_alani["Resim"] = df_calisma_alani["Resim"].str.replace(r"\.jpeg.*$", "", regex=True) + ".jpeg"
links = df_calisma_alani["Resim"].tolist()

df_calisma_alani = df_calisma_alani.fillna(0)
df_calisma_alani.replace([float('inf'), float('-inf')], 0, inplace=True)

#endregion

#region // AramaTerimleri Sütunundaki Tarihleri Ayıklama

date_pattern = r'(\d{1,2}\.\d{1,2}\.\d{4})'
df_calisma_alani['AramaTerimleri'] = df_calisma_alani['AramaTerimleri'].apply(
    lambda x: re.search(date_pattern, str(x)).group(1) if re.search(date_pattern, str(x)) else None
)

#endregion

#region // Bazı Sütunların Adını Güncelleme

df_calisma_alani.rename(columns={"AramaTerimleri": "Resim Yüklenme Tarihi"}, inplace=True)
df_calisma_alani.rename(columns={"AlisFiyati": "Alış Fiyatı"}, inplace=True)
df_calisma_alani.rename(columns={"SatisFiyati": "Satış Fiyatı"}, inplace=True)
df_calisma_alani.rename(columns={"UrunAdi": "Ürün Adı"}, inplace=True)
df_calisma_alani.rename(columns={"N11Kodu": "Mevsim"}, inplace=True)

#endregion

#region // Sütunların Biçim Ayarları ve Diğer Ayarlamalar

with pd.ExcelWriter('sonuc_excel.xlsx', engine='xlsxwriter') as writer:
    df_calisma_alani.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    max_col_width = max(df_calisma_alani["Ürün Adı"].astype(str).apply(len).max(), len("Ürün Adı")) + 2
    worksheet.set_column(0, 0, max_col_width)

    narrow_columns = ["Alış Fiyatı", "Satış Fiyatı", "GMT Stok Adedi", "SİTA Stok Adedi"]
    for col_name in narrow_columns:
        col_idx = df_calisma_alani.columns.get_loc(col_name)
        worksheet.set_column(col_idx, col_idx, 10)

    for i in range(1, len(df_calisma_alani.columns)):
        if df_calisma_alani.columns[i] not in narrow_columns:
            worksheet.set_column(i, i, 15)

    worksheet.set_row(0, 50)
    worksheet.freeze_panes(1, 0)

    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#D3D3D3',
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
        'border': 1
    })
    currency_format = workbook.add_format({'num_format': '#,##0.00₺', 'align': 'center', 'valign': 'vcenter'})
    shaded_currency_format = workbook.add_format({'bg_color': '#D9D9D9', 'num_format': '#,##0.00₺', 'align': 'center', 'valign': 'vcenter'})
    var_format = workbook.add_format({'bg_color': '#ffb994'})

    for col_num, value in enumerate(df_calisma_alani.columns.values):
        worksheet.write(0, col_num, value, header_format)

    center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
    left_align_format = workbook.add_format({'align': 'left', 'valign': 'vcenter'})
    shaded_format = workbook.add_format({'bg_color': '#D9D9D9', 'align': 'center', 'valign': 'vcenter'})
    shaded_left_align_format = workbook.add_format({'bg_color': '#D9D9D9', 'align': 'left', 'valign': 'vcenter'})

    for row_num, row in enumerate(df_calisma_alani.itertuples(), start=1):
        for col_num, value in enumerate(row[1:]):
            col_name = df_calisma_alani.columns[col_num]
            try:
                if col_num == 0:
                    link = links[row_num - 1] if row_num - 1 < len(links) else ""
                    if isinstance(link, str) and link.startswith("http"):
                        if row_num % 2 == 1:
                            worksheet.write_url(row_num, col_num, link, string=value, cell_format=shaded_left_align_format)
                        else:
                            worksheet.write_url(row_num, col_num, link, string=value, cell_format=left_align_format)
                    else:
                        if row_num % 2 == 1:
                            worksheet.write(row_num, col_num, value, shaded_left_align_format)
                        else:
                            worksheet.write(row_num, col_num, value, left_align_format)

                elif col_name in ["Alış Fiyatı", "Satış Fiyatı"]:
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
                continue

    worksheet.autofilter(0, 0, 0, len(df_calisma_alani.columns) - 1)
    worksheet.set_zoom(90)

os.rename("sonuc_excel.xlsx", "Nirvana.xlsx")

#endregion

#region // Kar Yüzdesi Sütununu Hesaplama

dosya_adi = "Nirvana.xlsx"
workbook = openpyxl.load_workbook(dosya_adi)
kopya_sayfa_adi = "Sheet1"

if kopya_sayfa_adi in workbook.sheetnames:
    sheet = workbook[kopya_sayfa_adi]
    # Başlık hücrelerini okumak için ilk satırı tarar ve değer->kolon indeksini sözlüğe alır
    basliklar = {cell.value: cell.column for cell in sheet[1] if cell.value}

    # "Net Satış Tarihi ve Adedi" başlığının Excel sütun indeksini bulma
    stok_adedi_kolon = basliklar.get("Net Satış Tarihi ve Adedi")

    if stok_adedi_kolon:
        # Eskiden +1 idi, artık +3 yaparak 3 sütun sonrasına koyacağız
        yeni_sutun_index = stok_adedi_kolon + 3

        # Eski sütundaki değerleri (ve stilleri) birebir kopyalayarak yeni sütuna aktarıyor
        for row in range(1, sheet.max_row + 1):
            eski_hucre = sheet.cell(row=row, column=stok_adedi_kolon)
            yeni_hucre = sheet.cell(row=row, column=yeni_sutun_index)
            yeni_hucre.value = eski_hucre.value
            if eski_hucre.has_style:
                from copy import copy
                yeni_hucre._style = copy(eski_hucre._style)

        # Yeni sütunun ilk hücresine (başlık satırı) isim verelim
        sheet.cell(row=1, column=yeni_sutun_index).value = "Kar Yüzdesi"

        # "Satış Fiyatı" ve "Alış Fiyatı" sütunlarını da sözlükten alalım
        satis_fiyati_kolon = basliklar.get("Satış Fiyatı")
        alis_fiyati_kolon = basliklar.get("Alış Fiyatı")

        # Her satır için kâr yüzdesi formülü uygula
        if satis_fiyati_kolon and alis_fiyati_kolon:
            for row in range(2, sheet.max_row + 1):
                satis_fiyati = sheet.cell(row=row, column=satis_fiyati_kolon).value
                alis_fiyati = sheet.cell(row=row, column=alis_fiyati_kolon).value
                yeni_hucre = sheet.cell(row=row, column=yeni_sutun_index)

                if satis_fiyati and alis_fiyati:
                    try:
                        kar_yuzdesi = (satis_fiyati - alis_fiyati) / satis_fiyati
                        yeni_hucre.value = kar_yuzdesi
                        # Excel’de yüzde formatı (ör. %10.00 vb.) görünmesi için
                        yeni_hucre.number_format = "0.00%"
                    except ZeroDivisionError:
                        yeni_hucre.value = None

workbook.save(dosya_adi)
gc.collect()


#endregion

#region // Gereksiz Excel Dosyalarını Silme

dosyalar = ["GMT ve SİTA.xlsx", "Satış Raporu.xlsx"]
for dosya in dosyalar:
    if os.path.exists(dosya):
        os.remove(dosya)

#endregion

#region // Sütunlara Açıklama Ekleme

dosya_yolu = "Nirvana.xlsx"
workbook = load_workbook(dosya_yolu)
sheet = workbook.active

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
    "SİTA Stok Adedi": "Ürünün SİTA üzerinde ne kadar stok adedi olduğunu belirtir",
    "Net Satış Tarihi ve Adedi": "Ürünün tüm renklerinin ve tüm bedenlerinin aktif olduğu son günü belirler ve o gün kaç adet sattığını belirtir",
    "Kar Yüzdesi": "Ürünün kar yüzdesini belirtir",
    "Son Transfer Tarihi": "Ürünün son tranfer edildiği tarihi belirtir",
    "Son İndirim Tarihi": "Ürünün son indirim yapıldığı tarihi belirtir"
}

for cell in sheet[1]:
    if cell.value in columns_with_comments:
        comment = Comment(columns_with_comments[cell.value], "Açıklama", width=400, height=300)
        cell.comment = comment

workbook.save(dosya_yolu)
del workbook
gc.collect()

#endregion

#region // Sigara Ürünleri Markadan Tespit Etme

file_path = "Nirvana.xlsx"
workbook = openpyxl.load_workbook(file_path)
sheet = workbook["Sheet1"]

urun_adi_column = None
marka_column = None

for col_index, column in enumerate(sheet[1], start=1):
    if column.value == "Ürün Adı":
        urun_adi_column = col_index
    elif column.value == "Marka":
        marka_column = col_index

if urun_adi_column is None or marka_column is None:
    raise ValueError("'Ürün Adı' veya 'Marka' sütunu bulunamadı.")

for row in sheet.iter_rows(min_row=2):
    urun_adi_cell = row[urun_adi_column - 1]
    marka_cell = row[marka_column - 1]

    if marka_cell.value and isinstance(marka_cell.value, str) and "Sigara Ürün" in marka_cell.value:
        urun_adi_cell.fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")

sheet.delete_cols(marka_column)
workbook.save(file_path)

#endregion





#region // Kopya Sayfa Oluşturma

dosya_adi = "Nirvana.xlsx"
workbook = openpyxl.load_workbook(dosya_adi)

if "Sheet1" in workbook.sheetnames:
    sheet1 = workbook["Sheet1"]
    sheet_copy = workbook.copy_worksheet(sheet1)
    sheet_copy.title = "Sheet1_Copy"

workbook.save(dosya_adi)

#endregion

#region // Liste Fiyatını Hesaplama

dosya_adi = "Nirvana.xlsx"
sheet_adi = "Sheet1_Copy"
workbook = openpyxl.load_workbook(dosya_adi)
if sheet_adi in workbook.sheetnames:
    sheet = workbook[sheet_adi]
    basliklar = {cell.value: cell.column for cell in sheet[1] if cell.value}
    alis_fiyati_kolon = basliklar.get("Alış Fiyatı")
    kategori_kolon = basliklar.get("Kategori")
    yeni_kolon_index = sheet.max_column + 1

    sheet.cell(row=1, column=yeni_kolon_index).value = "ListeFiyati2"

    for row in range(2, sheet.max_row + 1):
        alis_fiyati = sheet.cell(row=row, column=alis_fiyati_kolon).value
        kategori = sheet.cell(row=row, column=kategori_kolon).value

        if alis_fiyati is not None:
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

            if isinstance(kategori, str) and any(cat in kategori for cat in ["Parfüm", "Gözlük", "Saat", "Kolye", "Küpe", "Bileklik", "Bilezik"]):
                result *= 1.20
            else:
                result *= 1.10

            sheet.cell(row=row, column=yeni_kolon_index).value = result

workbook.save(dosya_adi)

#endregion

#region // Satış Fiyatı Liste Fiyatının Altındaysa Alış Fiyatını Kırmızı Yapma

dosya_adi = "Nirvana.xlsx"
sheet_adi = "Sheet1_Copy"
workbook = openpyxl.load_workbook(dosya_adi)
if sheet_adi in workbook.sheetnames:
    sheet = workbook[sheet_adi]

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
                fark = liste_fiyati2 - satis_fiyati
                if fark > 7:
                    alis_fiyati_hucre.font = Font(color="FF0000")

workbook.save(dosya_adi)

#endregion

#region // ListeFiyati2 Sütununu Silme

dosya_adi = "Nirvana.xlsx"
sheet_adi = "Sheet1_Copy"
workbook = openpyxl.load_workbook(dosya_adi)

if sheet_adi in workbook.sheetnames:
    sheet = workbook[sheet_adi]

    basliklar = {cell.value: cell.column for cell in sheet[1] if cell.value is not None}
    liste_fiyati2_kolon = basliklar.get("ListeFiyati2")

    if liste_fiyati2_kolon:
        sheet.delete_cols(liste_fiyati2_kolon)

workbook.save(dosya_adi)

#endregion

#region // Belirli Sütunları Silme

dosya_adi = "Nirvana.xlsx"
workbook = openpyxl.load_workbook(dosya_adi)
kopya_sayfa_adi = "Sheet1_Copy"

if kopya_sayfa_adi in workbook.sheetnames:
    sheet = workbook[kopya_sayfa_adi]

    silinecek_sutunlar = [
        "Stok Adedi Site ve Vega",
        "Ortalama Görüntülenme Adedi",
        "Kaç Güne Biter Site ve Vega",
        "Satış Fiyatı"
    ]

    basliklar = {cell.value: cell.column for cell in sheet[1] if cell.value is not None}

    for sutun_adi in reversed(silinecek_sutunlar):
        if sutun_adi in basliklar:
            sheet.delete_cols(basliklar[sutun_adi])

workbook.save(dosya_adi)

#endregion

#region // Sütunları Gizleme

workbook = openpyxl.load_workbook(dosya_adi)
if kopya_sayfa_adi in workbook.sheetnames:
    sheet = workbook[kopya_sayfa_adi]

    gizlenecek_sutunlar = [
        "Resim Yüklenme Tarihi",
        "Kategori",
        "GMT Stok Adedi",
        "SİTA Stok Adedi",
        "Son Transfer Tarihi",
        "Son İndirim Tarihi"
    ]

    basliklar = {cell.value: cell.column for cell in sheet[1] if cell.value is not None}

    for sutun_adi in gizlenecek_sutunlar:
        if sutun_adi in basliklar:
            sheet.column_dimensions[get_column_letter(basliklar[sutun_adi])].hidden = True

workbook.save(dosya_adi)

#endregion

#region // Kar Yüzdesi Sütununu Görünür Yapma

workbook = openpyxl.load_workbook(dosya_adi)
if kopya_sayfa_adi in workbook.sheetnames:
    sheet = workbook[kopya_sayfa_adi]
    basliklar = {cell.value: cell.column for cell in sheet[1] if cell.value is not None}
    if "Kar Yüzdesi" in basliklar:
        kar_yuzdesi_kolon = basliklar["Kar Yüzdesi"]
        sheet.column_dimensions[get_column_letter(kar_yuzdesi_kolon)].hidden = False

workbook.save(dosya_adi)

#endregion

#region // Sütunlara Filtreleme Özelliği Ekleme

workbook = openpyxl.load_workbook(dosya_adi)
if kopya_sayfa_adi in workbook.sheetnames:
    sheet = workbook[kopya_sayfa_adi]
    max_row = sheet.max_row
    max_col = sheet.max_column
    filter_ref = f"A1:{openpyxl.utils.get_column_letter(max_col)}{max_row}"
    sheet.auto_filter.ref = filter_ref

workbook.save(dosya_adi)

#endregion





#region // Kopya Sayfa Oluşturma

dosya_adi = "Nirvana.xlsx"
workbook = openpyxl.load_workbook(dosya_adi)

if "Sheet1" in workbook.sheetnames:
    sheet1 = workbook["Sheet1"]
    sheet_copy = workbook.copy_worksheet(sheet1)
    sheet_copy.title = "Sheet2_Copy"

workbook.save(dosya_adi)

#endregion

#region // Stoğu Olmayanları Silme

import openpyxl

# 1) Nirvana.xlsx dosyasını yükle
dosya_adi = "Nirvana.xlsx"
workbook = openpyxl.load_workbook(dosya_adi)

# 2) "Sheet2_Copy" sayfasını seç
sheet = workbook["Sheet2_Copy"]

# 3) Ürün Adı, İnstagram Stok Adedi ve Resim kolonlarının hangi sütunlarda olduğunu bulalım
product_name_col = None
instagram_stock_col = None
resim_col = None

for col in range(1, sheet.max_column + 1):
    header_value = sheet.cell(row=1, column=col).value
    if header_value == "Ürün Adı":
        product_name_col = col
    elif header_value == "İnstagram Stok Adedi":
        instagram_stock_col = col
    elif header_value == "Resim":
        resim_col = col

# Gerekli kolonları bulamazsak hata verelim
if product_name_col is None:
    raise ValueError('"Ürün Adı" başlığını bulamadım. Lütfen başlığı kontrol edin.')
if instagram_stock_col is None:
    raise ValueError('"İnstagram Stok Adedi" başlığını bulamadım. Lütfen başlığı kontrol edin.')
if resim_col is None:
    raise ValueError('"Resim" başlığını bulamadım. Lütfen başlığı kontrol edin.')

# 4) Satırları sondan başlayarak silme işlemini yapalım
#    - İnstagram Stok Adedi <= 0 olanlar
#    - Resim == 0 olanlar
for row in range(sheet.max_row, 1, -1):
    stock_cell = sheet.cell(row=row, column=instagram_stock_col)
    stock_value = stock_cell.value
    
    resim_cell = sheet.cell(row=row, column=resim_col)
    resim_value = resim_cell.value
    
    # Eğer stok <= 0 veya resim hücresi 0 ise satırı sil
    if (stock_value is not None and stock_value <= 0) or (resim_value == 0):
        sheet.delete_rows(row, 1)

# 5) Kalan satırlarda "Ürün Adı" kolonundaki hyperlink'i kaldırıp,
#    "Resim" kolonundaki değeri yeni hyperlink olarak ekleyelim
#    Ancak Ürün Adı hücresinin stilini değiştirmeyelim.
for row in range(2, sheet.max_row + 1):  # 1. satır başlık, 2'den itibaren veri var
    product_name_cell = sheet.cell(row=row, column=product_name_col)
    resim_cell = sheet.cell(row=row, column=resim_col)
    
    # Mevcut bağlantıyı kaldır (stili bozmadan)
    product_name_cell.hyperlink = None
    
    # Yeni bağlantı ekle (eğer "Resim" kolonunda değer varsa)
    if resim_cell.value:
        product_name_cell.hyperlink = resim_cell.value
        # NOT: Burada product_name_cell.style = "Hyperlink" kullanmıyoruz ki mevcut stil değişmesin

# 6) Son olarak "Resim" kolonunu sil
sheet.delete_cols(resim_col, 1)

# 7) Değişiklikleri kaydet
workbook.save("Nirvana.xlsx")



#endregion

#region // Liste Fiyatını Hesaplama

dosya_adi = "Nirvana.xlsx"
sheet_adi = "Sheet2_Copy"

workbook = openpyxl.load_workbook(dosya_adi)
if sheet_adi in workbook.sheetnames:
    sheet = workbook[sheet_adi]
    basliklar = {cell.value: cell.column for cell in sheet[1] if cell.value}

    # "Kar Yüzdesi", "Alış Fiyatı" ve "Kategori" sütunlarının kolon indekslerini alıyoruz
    kar_yuzdesi_kolon = basliklar.get("Kar Yüzdesi")
    alis_fiyati_kolon = basliklar.get("Alış Fiyatı")
    kategori_kolon = basliklar.get("Kategori")

    if kar_yuzdesi_kolon and alis_fiyati_kolon and kategori_kolon:
        # Yeni sütunu, "Kar Yüzdesi" sütununun hemen sağ tarafına ekleyeceğiz
        yeni_kolon_index = kar_yuzdesi_kolon + 1

        # 1) Kar Yüzdesi sütunundaki hücreleri biçimleriyle birlikte yeni sütuna kopyalayın
        for row in range(1, sheet.max_row + 1):
            eski_hucre = sheet.cell(row=row, column=kar_yuzdesi_kolon)
            yeni_hucre = sheet.cell(row=row, column=yeni_kolon_index)

            # Değer ve stil kopyası
            yeni_hucre.value = eski_hucre.value
            if eski_hucre.has_style:
                yeni_hucre._style = copy(eski_hucre._style)

        # 2) Yeni sütunun başlığını "Liste Fiyatı" olarak değiştirelim
        sheet.cell(row=1, column=yeni_kolon_index).value = "Liste Fiyatı"

        # 3) Her satır için alış fiyatına göre "Liste Fiyatı" değeri hesapla ve biçimlendir
        for row in range(2, sheet.max_row + 1):
            alis_fiyati = sheet.cell(row=row, column=alis_fiyati_kolon).value
            kategori = sheet.cell(row=row, column=kategori_kolon).value
            yeni_hucre = sheet.cell(row=row, column=yeni_kolon_index)

            if alis_fiyati is not None:
                # Alış fiyatına göre temel tutar
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

                # Kategori takı/parfüm/aksesuar vs. ise ek çarpan
                if (
                    isinstance(kategori, str) and 
                    any(cat in kategori for cat in ["Parfüm", "Gözlük", "Saat", "Kolye", "Küpe", "Bileklik", "Bilezik"])
                ):
                    result *= 1.20
                else:
                    result *= 1.10

                # Son olarak tam sayıya yuvarlayıp 0.99 ekliyoruz
                result = int(round(result)) + 0.99

                # Yeni sütuna yazıp para formatı (₺) ekleyelim
                yeni_hucre.value = result
                yeni_hucre.number_format = '#,##0.00₺'

workbook.save(dosya_adi)

#endregion

#region // Satış Fiyatı Liste Fiyatının Altındaysa Alış Fiyatını Kırmızı Yapma

dosya_adi = "Nirvana.xlsx"
sheet_adi = "Sheet2_Copy"
workbook = openpyxl.load_workbook(dosya_adi)
if sheet_adi in workbook.sheetnames:
    sheet = workbook[sheet_adi]

    basliklar = {cell.value: cell.column for cell in sheet[1] if cell.value}
    liste_fiyati2_kolon = basliklar.get("Liste Fiyatı")
    satis_fiyati_kolon = basliklar.get("Satış Fiyatı")
    alis_fiyati_kolon = basliklar.get("Alış Fiyatı")

    if liste_fiyati2_kolon and satis_fiyati_kolon and alis_fiyati_kolon:
        for row in range(2, sheet.max_row + 1):
            liste_fiyati2 = sheet.cell(row=row, column=liste_fiyati2_kolon).value
            satis_fiyati = sheet.cell(row=row, column=satis_fiyati_kolon).value
            alis_fiyati_hucre = sheet.cell(row=row, column=alis_fiyati_kolon)

            if liste_fiyati2 is not None and satis_fiyati is not None:
                fark = liste_fiyati2 - satis_fiyati
                if fark > 7:
                    alis_fiyati_hucre.font = Font(color="FF0000")

workbook.save(dosya_adi)

#endregion

#region // Belirli Sütunları Silme

dosya_adi = "Nirvana.xlsx"
workbook = openpyxl.load_workbook(dosya_adi)
kopya_sayfa_adi = "Sheet2_Copy"

if kopya_sayfa_adi in workbook.sheetnames:
    sheet = workbook[kopya_sayfa_adi]

    silinecek_sutunlar = [
        "Stok Adedi Site ve Vega",
        "Kaç Güne Biter Site ve Vega"

    ]

    basliklar = {cell.value: cell.column for cell in sheet[1] if cell.value is not None}

    for sutun_adi in reversed(silinecek_sutunlar):
        if sutun_adi in basliklar:
            sheet.delete_cols(basliklar[sutun_adi])

workbook.save(dosya_adi)

#endregion

#region // Sütunları Gizleme

workbook = openpyxl.load_workbook(dosya_adi)
if kopya_sayfa_adi in workbook.sheetnames:
    sheet = workbook[kopya_sayfa_adi]

    gizlenecek_sutunlar = [
        "GMT Stok Adedi",
        "SİTA Stok Adedi"
    ]

    basliklar = {cell.value: cell.column for cell in sheet[1] if cell.value is not None}

    for sutun_adi in gizlenecek_sutunlar:
        if sutun_adi in basliklar:
            sheet.column_dimensions[get_column_letter(basliklar[sutun_adi])].hidden = True

workbook.save(dosya_adi)

#endregion

#region // Kar Yüzdesi Sütununu Görünür Yapma

dosya_adi = "Nirvana.xlsx"
kopya_sayfa_adi = "Sheet2_Copy"

workbook = openpyxl.load_workbook(dosya_adi)
if kopya_sayfa_adi in workbook.sheetnames:
    sheet = workbook[kopya_sayfa_adi]
    
    # Başlık satırını tarayarak başlık -> kolon index ilişkisini (basliklar) bulalım
    basliklar = {cell.value: cell.column for cell in sheet[1] if cell.value is not None}
    
    # "Kar Yüzdesi" kolonunu görünür (hidden = False) yap
    if "Kar Yüzdesi" in basliklar:
        kar_yuzdesi_kolon = basliklar["Kar Yüzdesi"]
        sheet.column_dimensions[get_column_letter(kar_yuzdesi_kolon)].hidden = False
    
    # "Mevsim" kolonunu görünür (hidden = False) yap
    if "Mevsim" in basliklar:
        mevsim_kolon = basliklar["Mevsim"]
        sheet.column_dimensions[get_column_letter(mevsim_kolon)].hidden = False

workbook.save(dosya_adi)

#endregion

#region // Sütunlara Filtreleme Özelliği Ekleme

workbook = openpyxl.load_workbook(dosya_adi)
if kopya_sayfa_adi in workbook.sheetnames:
    sheet = workbook[kopya_sayfa_adi]
    max_row = sheet.max_row
    max_col = sheet.max_column
    filter_ref = f"A1:{openpyxl.utils.get_column_letter(max_col)}{max_row}"
    sheet.auto_filter.ref = filter_ref

workbook.save(dosya_adi)

#endregion






#region // Geriye Kalan Düzenlemeler


dosya_adi = "Nirvana.xlsx"

###############################################################################
# 1) Excel'i yükleme
###############################################################################
workbook = openpyxl.load_workbook(dosya_adi)

###############################################################################
# Yardımcı fonksiyonlar
###############################################################################
def find_column_index(sheet, column_name):
    """
    Belirtilen sheet'teki 1. satır (başlıklar) içerisinde column_name'e eşit olan
    kolonun indeksini döndürür. Bulamazsa None döndürür.
    """
    for col in range(1, sheet.max_column + 1):
        value = sheet.cell(row=1, column=col).value
        if value == column_name:
            return col
    return None

def freeze_top_row(sheet):
    """
    Verilen sheet'te üst satırı dondurur (scroll edince başlık satırı sabit kalır).
    """
    sheet.freeze_panes = "A2"

###############################################################################
# 1) (Ortak) - "Resim" = 0 => "Resim Yüklenme Tarihi" = "Resimsiz Ürün"
#             Boş kalan "Resim Yüklenme Tarihi" => "Tarih Yok"
###############################################################################
def fill_resim_and_tarih(sheet):
    """
    - "Resim" kolonunda değer 0 ise "Resim Yüklenme Tarihi" hücresine "Resimsiz Ürün" yaz.
    - "Resim Yüklenme Tarihi" boş (None veya "") ise "Tarih Yok" yaz.
    """
    resim_col = find_column_index(sheet, "Resim")
    resim_tarih_col = find_column_index(sheet, "Resim Yüklenme Tarihi")
    if not resim_col or not resim_tarih_col:
        return  # İlgili kolonlar yoksa işlem yapmayalım

    # "Resim" = 0 => "Resim Yüklenme Tarihi" = "Resimsiz Ürün"
    for row in range(2, sheet.max_row + 1):
        resim_value = sheet.cell(row=row, column=resim_col).value
        if resim_value == 0:
            sheet.cell(row=row, column=resim_tarih_col).value = "Resimsiz Ürün"

    # "Resim Yüklenme Tarihi" boş ise => "Tarih Yok"
    for row in range(2, sheet.max_row + 1):
        tarih_value = sheet.cell(row=row, column=resim_tarih_col).value
        if tarih_value is None or tarih_value == "":
            sheet.cell(row=row, column=resim_tarih_col).value = "Tarih Yok"

###############################################################################
# 2) (Sadece Sheet1) - "Resim" kolonunu "Stoksuz Üründe Hareket Var mı?" olarak
#     yeniden adlandır ve aşağıdaki koşullarla doldur:
#       - eğer "İnstagram Stok Adedi" <= 0 ve (Günlük Ortalama Satış Adedi > 0 veya 
#         Dünün Satış Adedi > 0) => "Evet"
#       - eğer "İnstagram Stok Adedi" <= 0 ve (her ikisi de 0 veya None) => "Hayır"
#       - eğer "İnstagram Stok Adedi" > 0 => "Stok Var"
###############################################################################
def update_resim_column_sheet1(sheet):
    old_header_col = find_column_index(sheet, "Resim")
    insta_stok_col = find_column_index(sheet, "İnstagram Stok Adedi")
    gunluk_ort_satis_col = find_column_index(sheet, "Günlük Ortalama Satış Adedi")
    dunun_satis_col = find_column_index(sheet, "Dünün Satış Adedi")

    if not old_header_col:
        return  # "Resim" kolonunu bulamadıysak işleme devam etmiyoruz
    
    # Kolon başlığını değiştir
    sheet.cell(row=1, column=old_header_col).value = "Stoksuz Üründe Hareket Var mı?"

    # Hücreleri doldur
    for row in range(2, sheet.max_row + 1):
        stok_val = sheet.cell(row=row, column=insta_stok_col).value if insta_stok_col else None
        gunluk_ort_val = sheet.cell(row=row, column=gunluk_ort_satis_col).value if gunluk_ort_satis_col else None
        dunun_val = sheet.cell(row=row, column=dunun_satis_col).value if dunun_satis_col else None

        # Boş (None) değerleri 0 gibi ele alalım
        stok_val = stok_val if stok_val is not None else 0
        gunluk_ort_val = gunluk_ort_val if gunluk_ort_val is not None else 0
        dunun_val = dunun_val if dunun_val is not None else 0

        if stok_val > 0:
            sonuc = "Stok Var"
        else:
            # stok <= 0
            if (gunluk_ort_val > 0) or (dunun_val > 0):
                sonuc = "Evet"
            else:
                sonuc = "Hayır"

        sheet.cell(row=row, column=old_header_col).value = sonuc

###############################################################################
# 3) (Sadece Sheet1_Copy) - "Resim" kolonunu "Resimsiz Üründe Hareket Var mı?"
#     olarak yeniden adlandırıp aynı stok ve satış kontrol mantığıyla doldur
###############################################################################
def update_resim_column_sheet1_copy(sheet):
    old_header_col = find_column_index(sheet, "Resim")
    insta_stok_col = find_column_index(sheet, "İnstagram Stok Adedi")
    gunluk_ort_satis_col = find_column_index(sheet, "Günlük Ortalama Satış Adedi")
    dunun_satis_col = find_column_index(sheet, "Dünün Satış Adedi")

    if not old_header_col:
        return  # "Resim" kolonunu bulamadıysak işleme devam etmiyoruz
    
    # Kolon başlığını değiştir
    sheet.cell(row=1, column=old_header_col).value = "Stoksuz Üründe Hareket Var mı?"

    # Hücreleri doldur
    for row in range(2, sheet.max_row + 1):
        stok_val = sheet.cell(row=row, column=insta_stok_col).value if insta_stok_col else None
        gunluk_ort_val = sheet.cell(row=row, column=gunluk_ort_satis_col).value if gunluk_ort_satis_col else None
        dunun_val = sheet.cell(row=row, column=dunun_satis_col).value if dunun_satis_col else None

        # Boş (None) değerleri 0 gibi ele alalım
        stok_val = stok_val if stok_val is not None else 0
        gunluk_ort_val = gunluk_ort_val if gunluk_ort_val is not None else 0
        dunun_val = dunun_val if dunun_val is not None else 0

        if stok_val > 0:
            sonuc = "Stok Var"
        else:
            # stok <= 0
            if (gunluk_ort_val > 0) or (dunun_val > 0):
                sonuc = "Evet"
            else:
                sonuc = "Hayır"

        sheet.cell(row=row, column=old_header_col).value = sonuc

###############################################################################
# 4) (Sheet1 & Sheet1_Copy) - "Kaç Güne Biter Her Şey Dahil" kolonunda,
#    "Stok Adedi Her Şey Dahil" <= 0 => "Stok Yok"
###############################################################################
def fill_stok_yok_her_sey_dahil(sheet):
    kac_gune_biter_col = find_column_index(sheet, "Kaç Güne Biter Her Şey Dahil")
    stok_hersey_col = find_column_index(sheet, "Stok Adedi Her Şey Dahil")
    if not (kac_gune_biter_col and stok_hersey_col):
        return

    for row in range(2, sheet.max_row + 1):
        stok_val = sheet.cell(row=row, column=stok_hersey_col).value
        if stok_val is not None and stok_val <= 0:
            sheet.cell(row=row, column=kac_gune_biter_col).value = "Stok Yok"

###############################################################################
# 5) (Sadece Sheet1) - "Kaç Güne Biter Site ve Vega" kolonunda,
#    "Stok Adedi Site ve Vega" <= 0 => "Stok Yok"
###############################################################################
def fill_stok_yok_site_vega(sheet):
    kac_gune_site_vega_col = find_column_index(sheet, "Kaç Güne Biter Site ve Vega")
    stok_site_vega_col = find_column_index(sheet, "Stok Adedi Site ve Vega")
    if not (kac_gune_site_vega_col and stok_site_vega_col):
        return

    for row in range(2, sheet.max_row + 1):
        stok_val = sheet.cell(row=row, column=stok_site_vega_col).value
        if stok_val is not None and stok_val <= 0:
            sheet.cell(row=row, column=kac_gune_site_vega_col).value = "Stok Yok"

###############################################################################
# 6) (Sheet1 & Sheet1_Copy) - "Kategori" kolonunda veri 0 ise => "Kategori Yok"
###############################################################################
def fill_kategori_yok(sheet):
    kategori_col = find_column_index(sheet, "Kategori")
    if not kategori_col:
        return

    for row in range(2, sheet.max_row + 1):
        val = sheet.cell(row=row, column=kategori_col).value
        if val == 0:
            sheet.cell(row=row, column=kategori_col).value = "Kategori Yok"

###############################################################################
# 7) (Sheet1_Copy & Sheet2_Copy) - İlk satırı dondur (freeze panes)
###############################################################################

###############################################################################
# Uygulamaya geçiş
###############################################################################

# A) Sheet1 - 1. maddedeki "Resim" ve "Resim Yüklenme Tarihi" doldurma:
if "Sheet1" in workbook.sheetnames:
    fill_resim_and_tarih(workbook["Sheet1"])         # (Madde 1)
    update_resim_column_sheet1(workbook["Sheet1"])   # (Madde 2)
    fill_stok_yok_her_sey_dahil(workbook["Sheet1"])  # (Madde 4)
    fill_stok_yok_site_vega(workbook["Sheet1"])      # (Madde 5)
    fill_kategori_yok(workbook["Sheet1"])            # (Madde 6)

# B) Sheet1_Copy
if "Sheet1_Copy" in workbook.sheetnames:
    fill_resim_and_tarih(workbook["Sheet1_Copy"])    # (Madde 1 tekrarı)
    update_resim_column_sheet1_copy(workbook["Sheet1_Copy"])  # (Madde 3)
    fill_stok_yok_her_sey_dahil(workbook["Sheet1_Copy"])       # (Madde 4)
    fill_kategori_yok(workbook["Sheet1_Copy"])                 # (Madde 6)
    freeze_top_row(workbook["Sheet1_Copy"])                    # (Madde 7 - freeze)

# C) Sheet2_Copy - sadece 7. madde gereği freeze topl row
if "Sheet2_Copy" in workbook.sheetnames:
    freeze_top_row(workbook["Sheet2_Copy"])

# D) Son olarak kaydet
workbook.save(dosya_adi)


#endregion

#region // Sayfaların İsmini Değiştirme

import openpyxl

dosya_adi = "Nirvana.xlsx"

# Eski sayfa adları -> Yeni sayfa adları
sayfa_adi_haritasi = {
    "Sheet1": "Genel Rapor",
    "Sheet1_Copy": "RPT Raporu",
    "Sheet2_Copy": "İndirim Raporu"
}

# Excel dosyasını aç
workbook = openpyxl.load_workbook(dosya_adi)

# Her bir eşleşmeyi kontrol ederek isim değiştirme
for eski_ad, yeni_ad in sayfa_adi_haritasi.items():
    if eski_ad in workbook.sheetnames:
        # Aynı yeni isim kullanılıyorsa çakışmayı önlemek için kontrol edebilirsiniz
        if yeni_ad in workbook.sheetnames:
            print(f"'{yeni_ad}' isimli bir sayfa zaten var. Lütfen farklı bir isim kullanın.")
        else:
            workbook[eski_ad].title = yeni_ad
            
    else:
        print(f"'{eski_ad}' isimli bir sayfa bulunamadı, atlanıyor.")

# Değişiklikleri kaydet
workbook.save(dosya_adi)

#endregion
