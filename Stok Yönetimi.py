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
from copy import copy
from openpyxl.worksheet.table import Table, TableStyleInfo


init(autoreset=True)
warnings.filterwarnings("ignore")
colorama.init(autoreset=True)

#endregion

#region // Entegrasyondan Önce mi Sonra mı Kontrolü

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

#endregion





# ------------------------------------------------------------
# Yardımcı Fonksiyonlar
# ------------------------------------------------------------

def parse_decimal_string(value):
    """
    Örneğin "1,0000" veya "99,9900" gibi virgüllü ondalıkları float'a çevirir.
    "129,9900" -> 129.99
    """
    s = str(value).strip()      # string'e çevir, boşlukları temizle
    s = s.replace(',', '.')     # Virgülleri noktaya çevir
    try:
        return float(s)
    except ValueError:
        return 0.0

def extract_product_code(urun_adi):
    """
    Örnek: "TEAM BRODE - 62532. Kırmızı" -> "62532"
    '- (\\d+)\\.' kalıbını arar; yoksa None döner.
    """
    match = re.search(r' - (\d+)\.', urun_adi)
    return match.group(1) if match else None

def extract_color(urun_adi):
    """
    Örnek: "TEAM BRODE - 62532. Kırmızı" -> "Kırmızı"
    Yöntem: ' - ' bazında böldükten sonra ilk parçanın son kelimesini alıyoruz.
    Ama veriniz farklı yapıdaysa lütfen uyarlayın.
    """
    parts = re.split(r' - ', urun_adi)
    if len(parts) > 0:
        before_part = parts[0].strip()
        words = before_part.split()
        if words:
            return words[-1]
    return None

# ------------------------------------------------------------
# 1) Supabase üzerinden Ürün Listesi (tum_urun_listesi.xlsx) İndirme ve İşleme
# ------------------------------------------------------------

SUPABASE_URL = "https://zmvsatlvobhdaxxgtoap.supabase.co"
SUPABASE_KEY = (
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9."
    "eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InptdnNhdGx2b2JoZGF4eGd0b2FwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDAxNzIxMzksImV4cCI6MjA1NTc0ODEzOX0."
    "lJLudSfixMbEOkJmfv22MsRLofP7ZjFkbGj26xF3dts"
)

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# "tum_urun_listesi.xlsx" indirilir (Belleğe)
response = supabase.storage.from_("tum_urun_listesi").download("tum_urun_listesi.xlsx")
temp_df = pd.read_excel(BytesIO(response))

# Kullanmak istediğimiz sütunlar
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
df = temp_df[selected_columns].copy()

# Benzersiz kayıtlar (StokAdedi hariç)
unique_cols = [c for c in selected_columns if c != "StokAdedi"]
unique_df = df.drop(columns=["StokAdedi"]).drop_duplicates(subset=unique_cols)

# UrunAdi bazında StokAdedi toplanır
stokadedi_sums = df.groupby("UrunAdi")["StokAdedi"].sum().reset_index()

# Birleştir => final_df
final_df = pd.merge(unique_df, stokadedi_sums, on="UrunAdi", how="left")

# ------------------------------------------------------------
# 2) Satış Raporu İndirme (Selenium) + Bellekte Parse Edip final_df'ye Ekleme
# ------------------------------------------------------------

chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--incognito")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

service = Service()
driver = webdriver.Chrome(service=service, options=chrome_options)

username = "mustafa_kod@haydigiy.com"
password = "123456"
login_url = "https://www.siparis.haydigiy.com/kullanici-giris/?ReturnUrl=%2Fadmin"
desired_page_url = "https://www.siparis.haydigiy.com/admin/exportorder/edit/154/"

try:
    # 2.1) Giriş
    driver.get(login_url)
    time.sleep(2)
    driver.find_element(By.NAME, "EmailOrPhone").send_keys(username)
    driver.find_element(By.NAME, "Password").send_keys(password)
    driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
    time.sleep(3)

    # 2.2) İlgili sayfaya gidip tarih ayarlarını girmek
    driver.get(desired_page_url)
    time.sleep(2)

    yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
    formatted_date_no_leading = f"{yesterday.day}.{yesterday.month}.{yesterday.year}"

    end_date_input = driver.find_element(By.ID, "EndDate")
    end_date_input.clear()
    end_date_input.send_keys(formatted_date_no_leading)

    start_date_input = driver.find_element(By.ID, "StartDate")
    start_date_input.clear()
    start_date_input.send_keys(formatted_date_no_leading)

    save_button = driver.find_element(By.CSS_SELECTOR, 'button.btn.btn-primary[name="save"]')
    save_button.click()

except Exception as e:
    print(colorama.Fore.RED + f"Giriş veya tarih ayarı sırasında hata: {e}" + colorama.Style.RESET_ALL)

finally:
    driver.quit()

# 2.3) Bellekte Satış Raporu verisini alalım
url = "https://www.siparis.haydigiy.com/FaprikaOrderXls/GZPCKE/1/"
resp = requests.get(url)
sales_df = pd.read_excel(BytesIO(resp.content))

# 2.4) "Adet" ve "ToplamFiyat" float'a çevirelim
columns_to_keep = ["UrunAdi", "Adet", "ToplamFiyat"]
sales_df = sales_df[columns_to_keep]
sales_df["Adet"] = sales_df["Adet"].apply(parse_decimal_string)
sales_df["ToplamFiyat"] = sales_df["ToplamFiyat"].apply(parse_decimal_string)

# 2.5) UrunAdi bazında sum
sales_sums = sales_df.groupby("UrunAdi", as_index=False).agg({
    "Adet": "sum",
    "ToplamFiyat": "sum"
})

# 2.6) final_df ile birleştir => final_merged_df
final_merged_df = pd.merge(final_df, sales_sums, on="UrunAdi", how="left")
final_merged_df["Adet"] = final_merged_df["Adet"].fillna(0)
final_merged_df["ToplamFiyat"] = final_merged_df["ToplamFiyat"].fillna(0)

# ------------------------------------------------------------
# 3) "UrunAdi Duzenleme" ve "UrunAdi ve Renk" Kolonlarını Oluşturma
# ------------------------------------------------------------

final_merged_df["UrunAdi Duzenleme"] = final_merged_df["UrunAdi"].apply(extract_product_code)
final_merged_df["UrunAdi Duzenleme"] = final_merged_df["UrunAdi Duzenleme"].fillna("").astype(str)

def build_urun_adi_ve_renk(row):
    """
    "UrunAdi ve Renk" = UrunAdi Duzenleme + " - " + Renk
    """
    product_code_str = row["UrunAdi Duzenleme"]
    color_str = extract_color(row["UrunAdi"])  # Örnek: "Kırmızı"
    if not color_str:
        color_str = ""
    return product_code_str + " - " + color_str

final_merged_df["UrunAdi ve Renk"] = final_merged_df.apply(build_urun_adi_ve_renk, axis=1)

# ------------------------------------------------------------
# YENİ EKLEME: "urunyonetimi" Tablosundan "id, urunkodu, renk, fiyat" Kolonlarını Çekip
# ID En Büyük Olan Kaydı Baz Alarak "Satın Alma Fiyatı" Sütununu Oluşturma
# ------------------------------------------------------------

# 1) Supabase'den veriyi alıyoruz
response_data_fiyat = supabase.table("urunyonetimi").select("id, urunkodu, renk, alisfiyati").execute()
df_urunyonetimi_fiyat = pd.DataFrame(response_data_fiyat.data)

# 2) urunkodu ve renk'i birleştireceğimiz kolon
df_urunyonetimi_fiyat["urunkodu_renk"] = df_urunyonetimi_fiyat["urunkodu"].astype(str) + " - " + df_urunyonetimi_fiyat["renk"].astype(str)

# 3) Aynı urunkodu_renk'e sahip birden fazla satır varsa, ID değeri en büyük olanı al
#    Bunun için önce ID'ye göre azalan sıralama yapar, sonra drop_duplicates ile en üsttekini tutarız.
df_urunyonetimi_fiyat.sort_values("id", ascending=False, inplace=True)
df_urunyonetimi_fiyat.drop_duplicates(subset=["urunkodu_renk"], keep="first", inplace=True)

# 4) final_merged_df ile merge
final_merged_df = final_merged_df.merge(
    df_urunyonetimi_fiyat[["urunkodu_renk", "alisfiyati"]],
    how="left",
    left_on="UrunAdi ve Renk",
    right_on="urunkodu_renk"
)

# 5) Gelen "fiyat" kolonunu "Satın Alma Fiyatı" olarak yeniden adlandırıyoruz
final_merged_df.rename(columns={"alisfiyati": "Son Satın Alma Fiyatı"}, inplace=True)

# 6) "urunkodu_renk" geçici kolonunu atabiliriz
final_merged_df.drop("urunkodu_renk", axis=1, inplace=True)

# 7) "Satın Alma Fiyatı" kolonunu "AlisFiyati" kolonunun hemen sağına yerleştirelim
cols = list(final_merged_df.columns)
if "AlisFiyati" in cols and "Son Satın Alma Fiyatı" in cols:
    alis_idx = cols.index("AlisFiyati")
    satinalma_idx = cols.index("Son Satın Alma Fiyatı")
    col_to_move = cols.pop(satinalma_idx)
    cols.insert(alis_idx + 1, col_to_move)
    final_merged_df = final_merged_df[cols]


# ------------------------------------------------------------
# 4) GMT ve SİTA Verilerini Çekme (Supabase tablosu "urunyonetimi")
# ------------------------------------------------------------

all_data = []
start = 0
page_size = 1000

while True:
    end = start + page_size - 1
    response_data = (
        supabase.table("urunyonetimi")
        .select("urunkodu, renk, acilmamisadet, gmtsitalabel")
        .in_("gmtsitalabel", ["GMT", "SİTA", "Yarım GMT"])
        .gt("acilmamisadet", 0)
        .range(start, end)
        .execute()
    )
    data = response_data.data
    if not data:
        break
    all_data.extend(data)
    start += page_size

df_gmtsita = pd.DataFrame(all_data)

# Renk sütununu düzenleyelim (İlk harf büyük)
df_gmtsita["renk"] = df_gmtsita["renk"].apply(lambda x: x.capitalize() if isinstance(x, str) else x)

# GMT + Yarım GMT
df_gmt = df_gmtsita[df_gmtsita["gmtsitalabel"].isin(["GMT", "Yarım GMT"])]
df_gmt_grouped = df_gmt.groupby(["urunkodu", "renk"], as_index=False)["acilmamisadet"].sum()

df_gmt_final = pd.DataFrame()
df_gmt_final["GMT Ürün Kodu"] = pd.to_numeric(df_gmt_grouped["urunkodu"], errors="coerce").astype("Int64")
df_gmt_final["GMT Ürün Adı"] = df_gmt_grouped["urunkodu"].astype(str) + " - " + df_gmt_grouped["renk"]
df_gmt_final["GMT Stok Adedi"] = df_gmt_grouped["acilmamisadet"]

# SİTA
df_sita = df_gmtsita[df_gmtsita["gmtsitalabel"] == "SİTA"]
df_sita_grouped = df_sita.groupby(["urunkodu", "renk"], as_index=False)["acilmamisadet"].sum()

df_sita_final = pd.DataFrame()
df_sita_final["SİTA Ürün Kodu"] = pd.to_numeric(df_sita_grouped["urunkodu"], errors="coerce").astype("Int64")
df_sita_final["SİTA Ürün Adı"] = df_sita_grouped["urunkodu"].astype(str) + " - " + df_sita_grouped["renk"]
df_sita_final["SİTA Stok Adedi"] = df_sita_grouped["acilmamisadet"]

# ------------------------------------------------------------
# 5) GMT ve SİTA Verilerini Ana Tabloya Çektirme (Etopla Mantığı)
# ------------------------------------------------------------

df_calisma_alani = final_merged_df.copy()

used_gmt_indices_step1 = []
used_sita_indices_step1 = []

# 5.1) "UrunAdi ve Renk" -> "GMT Ürün Adı" / "SİTA Ürün Adı"
for idx, row in df_calisma_alani.iterrows():
    urun_adi_ve_renk = row.get("UrunAdi ve Renk", "")

    # GMT eşleştirme
    matching_gmt = df_gmt_final[df_gmt_final["GMT Ürün Adı"] == urun_adi_ve_renk]
    if not matching_gmt.empty:
        matched_index = matching_gmt.index[0]
        df_calisma_alani.at[idx, "GMT Stok Adedi"] = matching_gmt.iloc[0]["GMT Stok Adedi"]
        used_gmt_indices_step1.append(matched_index)
    else:
        df_calisma_alani.at[idx, "GMT Stok Adedi"] = None

    # SİTA eşleştirme
    matching_sita = df_sita_final[df_sita_final["SİTA Ürün Adı"] == urun_adi_ve_renk]
    if not matching_sita.empty:
        matched_index = matching_sita.index[0]
        df_calisma_alani.at[idx, "SİTA Stok Adedi"] = matching_sita.iloc[0]["SİTA Stok Adedi"]
        used_sita_indices_step1.append(matched_index)
    else:
        df_calisma_alani.at[idx, "SİTA Stok Adedi"] = None

df_gmt_final = df_gmt_final.drop(used_gmt_indices_step1).reset_index(drop=True)
df_sita_final = df_sita_final.drop(used_sita_indices_step1).reset_index(drop=True)

used_gmt_indices_step2 = []
used_sita_indices_step2 = []

# 5.2) "UrunAdi Duzenleme" -> "GMT Ürün Kodu" / "SİTA Ürün Kodu"
for idx, row in df_calisma_alani.iterrows():
    urun_kodu = row.get("UrunAdi Duzenleme", "")

    # GMT kontrol
    current_gmt_stok = row.get("GMT Stok Adedi", 0)
    if pd.isna(current_gmt_stok) or current_gmt_stok == 0:
        matching_gmt_code = df_gmt_final[df_gmt_final["GMT Ürün Kodu"].astype(str) == urun_kodu]
        if not matching_gmt_code.empty:
            matched_index = matching_gmt_code.index[0]
            gmt_stok = matching_gmt_code.iloc[0]["GMT Stok Adedi"]
            df_calisma_alani.at[idx, "GMT Stok Adedi"] = "GMT'de Var" if gmt_stok > 0 else gmt_stok
            used_gmt_indices_step2.append(matched_index)

    # SİTA kontrol
    current_sita_stok = row.get("SİTA Stok Adedi", 0)
    if pd.isna(current_sita_stok) or current_sita_stok == 0:
        matching_sita_code = df_sita_final[df_sita_final["SİTA Ürün Kodu"].astype(str) == urun_kodu]
        if not matching_sita_code.empty:
            matched_index = matching_sita_code.index[0]
            sita_stok = matching_sita_code.iloc[0]["SİTA Stok Adedi"]
            df_calisma_alani.at[idx, "SİTA Stok Adedi"] = "SİTA'da Var" if sita_stok > 0 else sita_stok
            used_sita_indices_step2.append(matched_index)

df_gmt_final = df_gmt_final.drop(used_gmt_indices_step2).reset_index(drop=True)
df_sita_final = df_sita_final.drop(used_sita_indices_step2).reset_index(drop=True)

# ------------------------------------------------------------
# 6) Tek Çıktı: "Nirvana.xlsx"
# ------------------------------------------------------------

df_calisma_alani.to_excel("Nirvana.xlsx", index=False)























# ------------------------------------------------------------
# EK ADIMLAR
# ------------------------------------------------------------
# 1) "ListeFiyatı" sütunu oluşturma (openpyxl ile)
#    Bu aşamada 'Kar Yüzdesi', 'Alış Fiyatı', 'Kategori' kolonlarının
#    zaten var olduğunu varsayıyoruz.



dosya_adi = "Nirvana.xlsx"
workbook = openpyxl.load_workbook(dosya_adi)
kopya_sayfa_adi = "Sheet1"  # Kendi sayfa adınıza göre değiştirebilirsiniz

if kopya_sayfa_adi in workbook.sheetnames:
    sheet = workbook[kopya_sayfa_adi]
    
    # Başlıkları okuyup (ilk satır) sütun isimlerini ve indekslerini sözlüğe atıyoruz
    basliklar = {cell.value: cell.column for cell in sheet[1] if cell.value}

    # "VaryasyonGittiGidiyorKodu" sütununu bul
    vgk_kolon = basliklar.get("VaryasyonGittiGidiyorKodu")

    if vgk_kolon:
        # 1) En sağdaki ilk boş sütunu belirleyelim
        yeni_sutun_index = sheet.max_column + 1  # max_column’ın bir sonrasına eklenecek

        # 2) Yeni sütuna, "VaryasyonGittiGidiyorKodu" değerlerini kopyalayalım (opsiyonel)
        for row in range(1, sheet.max_row + 1):
            eski_hucre = sheet.cell(row=row, column=vgk_kolon)
            yeni_hucre = sheet.cell(row=row, column=yeni_sutun_index)
            yeni_hucre.value = eski_hucre.value
            # Stil kopyalama (opsiyonel)
            if eski_hucre.has_style:
                yeni_hucre._style = copy(eski_hucre._style)

        # 3) Yeni sütunun başlığını "Kar Yüzdesi" yapalım (1. satır)
        sheet.cell(row=1, column=yeni_sutun_index).value = "Kar Yüzdesi"

        # 4) Satış Fiyatı ve Alış Fiyatı sütunlarının indekslerini alalım
        satis_fiyati_kolon = basliklar.get("SatisFiyati")
        alis_fiyati_kolon = basliklar.get("AlisFiyati")

        # 5) Her satır için kâr yüzdesi hesapla
        #    (Kar yüzdesi = (Satış - Alış) / Satış) * 100
        if satis_fiyati_kolon and alis_fiyati_kolon:
            for row in range(2, sheet.max_row + 1):
                satis_fiyati = sheet.cell(row=row, column=satis_fiyati_kolon).value
                alis_fiyati = sheet.cell(row=row, column=alis_fiyati_kolon).value

                # "Kar Yüzdesi" hücresi (yeni sütun)
                kar_yuzdesi_hucre = sheet.cell(row=row, column=yeni_sutun_index)

                if (satis_fiyati is not None) and (alis_fiyati is not None):
                    try:
                        kar_yuzdesi = (satis_fiyati - alis_fiyati) / satis_fiyati * 100
                        kar_yuzdesi_hucre.value = kar_yuzdesi
                        kar_yuzdesi_hucre.number_format = "0.00"
                    except ZeroDivisionError:
                        kar_yuzdesi_hucre.value = None

        # Değişiklikleri kaydedip workbook'u kapatalım
        workbook.save(dosya_adi)
        gc.collect()


# ============================================================
# 2) "ListeFiyatı" kolonunu oluşturma
# ============================================================
workbook = openpyxl.load_workbook(dosya_adi)

# Açılacak sayfa ismi ("Sheet1" veya oluşturulan tablo ismi olabilir).
sheet_adi = workbook.active.title  # Burada direkt aktif sayfanın adını aldık
# İsterseniz sheet_adi = "Sheet1" şeklinde de doğrudan belirtebilirsiniz.

if sheet_adi in workbook.sheetnames:
    sheet = workbook[sheet_adi]

    # Başlık hücrelerini (ilk satır) bir sözlükte tutuyoruz
    basliklar = {cell.value: cell.column for cell in sheet[1] if cell.value}

    kar_yuzdesi_kolon = basliklar.get("Kar Yüzdesi")
    alis_fiyati_kolon = basliklar.get("AlisFiyati")
    kategori_kolon = basliklar.get("Kategori")

    if kar_yuzdesi_kolon and alis_fiyati_kolon and kategori_kolon:
        # Yeni sütunu, "Kar Yüzdesi" sütununun hemen sağ tarafına ekleyeceğiz
        yeni_kolon_index = kar_yuzdesi_kolon + 1

        # 1) Kar Yüzdesi sütunundaki hücreleri biçimiyle birlikte yeni sütuna kopyalayalım
        for row in range(1, sheet.max_row + 1):
            eski_hucre = sheet.cell(row=row, column=kar_yuzdesi_kolon)
            yeni_hucre = sheet.cell(row=row, column=yeni_kolon_index)
            yeni_hucre.value = eski_hucre.value
            if eski_hucre.has_style:
                yeni_hucre._style = copy(eski_hucre._style)

        # 2) Yeni sütunun başlığını "ListeFiyatı" yap
        sheet.cell(row=1, column=yeni_kolon_index).value = "ListeFiyatı"

        # 3) Her satır için alış fiyatına göre "Liste Fiyatı" hesapla
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

                # Kategori takı/parfüm/aksesuar vs. ise ekstra çarpan
                if (
                    isinstance(kategori, str) and 
                    any(cat in kategori for cat in ["Parfüm", "Gözlük", "Saat", "Kolye", "Küpe", "Bileklik", "Bilezik"])
                ):
                    result *= 1.20
                else:
                    result *= 1.10

                # Son olarak tam sayıya yuvarlayıp 0.99 ekleyelim
                result = int(round(result)) + 0.99

                # Para formatı (₺) ekle
                yeni_hucre.value = result
                yeni_hucre.number_format = '#,##0.00₺'

workbook.save(dosya_adi)
gc.collect()


# ------------------------------------------------------------
# 2) "AramaTerimleri" kolonu için tarih ayıklama
# ------------------------------------------------------------
df_calisma_alani = pd.read_excel("Nirvana.xlsx")

date_pattern = r'(\d{1,2}\.\d{1,2}\.\d{4})'
df_calisma_alani['AramaTerimleri'] = df_calisma_alani['AramaTerimleri'].apply(
    lambda x: re.search(date_pattern, str(x)).group(1) if re.search(date_pattern, str(x)) else None
)

# ------------------------------------------------------------
# 3) "Kategori" kolonunda düzenleme
# ------------------------------------------------------------
df_calisma_alani['Kategori'] = df_calisma_alani['Kategori'].fillna("")

def extract_category(text):
    if not isinstance(text, str):
        return None
    match = re.search(r'>\s*([^;]+)', text)
    if match:
        return match.group(1).strip()
    elif "TESETTÜR" in text:
        return "TESETTÜR"
    return None

df_calisma_alani['Kategori'] = df_calisma_alani['Kategori'].apply(extract_category)

# ------------------------------------------------------------
# 4) Yeni kolonlar: "Stok Adedi Her Şey Dahil" ve "Stok Adedi Site ve Vega"
# ------------------------------------------------------------
# Burada kullanılan kolon adlarına dikkat: 
# "GMT Stok Adedi", "SİTA Stok Adedi", "StokAdedi" ve "VaryasyonMorhipoKodu" 
# yoksa oluşturulabilir.

if "GMT Stok Adedi" not in df_calisma_alani.columns:
    df_calisma_alani["GMT Stok Adedi"] = 0
if "SİTA Stok Adedi" not in df_calisma_alani.columns:
    df_calisma_alani["SİTA Stok Adedi"] = 0
if "StokAdedi" not in df_calisma_alani.columns:
    df_calisma_alani["StokAdedi"] = 0
if "VaryasyonMorhipoKodu" not in df_calisma_alani.columns:
    df_calisma_alani["VaryasyonMorhipoKodu"] = 0

gmt_numeric = pd.to_numeric(df_calisma_alani["GMT Stok Adedi"], errors="coerce").fillna(0)
sita_numeric = pd.to_numeric(df_calisma_alani["SİTA Stok Adedi"], errors="coerce").fillna(0)
stok_adedi_numeric = pd.to_numeric(df_calisma_alani["StokAdedi"], errors="coerce").fillna(0)
n11_zimmet_numeric = pd.to_numeric(df_calisma_alani["VaryasyonMorhipoKodu"], errors="coerce").fillna(0)

df_calisma_alani["Stok Adedi Her Şey Dahil"] = stok_adedi_numeric + n11_zimmet_numeric + gmt_numeric + sita_numeric
df_calisma_alani["Stok Adedi Site ve Vega"] = stok_adedi_numeric + n11_zimmet_numeric

df_calisma_alani['StokAdedi'].fillna(0, inplace=True)
df_calisma_alani['VaryasyonMorhipoKodu'].fillna(0, inplace=True)
df_calisma_alani['GMT Stok Adedi'].fillna(0, inplace=True)
df_calisma_alani['SİTA Stok Adedi'].fillna(0, inplace=True)

# ------------------------------------------------------------
# 5) Yeni kolonlar: "Kaç Güne Biter Her Şey Dahil" ve "Kaç Güne Biter Site ve Vega"
#    Hesaplama: Stok / MorhipoKodu
# ------------------------------------------------------------
# Kolon isimleri tutarlılık için kontrol ediliyor
if "MorhipoKodu" not in df_calisma_alani.columns:
    df_calisma_alani["MorhipoKodu"] = 0

df_calisma_alani['MorhipoKodu'].fillna(0, inplace=True)
df_calisma_alani["Kaç Güne Biter Her Şey Dahil"] = "Satış Adedi Yok"
df_calisma_alani["Kaç Güne Biter Site ve Vega"] = "Satış Adedi Yok"

non_zero_mask = df_calisma_alani["MorhipoKodu"] != 0
df_calisma_alani.loc[non_zero_mask, "Kaç Güne Biter Her Şey Dahil"] = round(
    df_calisma_alani["Stok Adedi Her Şey Dahil"] / df_calisma_alani["MorhipoKodu"]
)

df_calisma_alani.loc[non_zero_mask, "Kaç Güne Biter Site ve Vega"] = round(
    df_calisma_alani["Stok Adedi Site ve Vega"] / df_calisma_alani["MorhipoKodu"]
)

# ------------------------------------------------------------
# 6) "Görüntülenmenin Satışa Dönüş Oranı" kolonu
#    "HepsiBuradaKodu" -> "Ortalama Görüntülenme Adedi" rename
# ------------------------------------------------------------
# Kolon yoksa ekliyoruz
if "HepsiBuradaKodu" not in df_calisma_alani.columns:
    df_calisma_alani["HepsiBuradaKodu"] = 0

df_calisma_alani = df_calisma_alani.rename(columns={"HepsiBuradaKodu": "Ortalama Görüntülenme Adedi"})
df_calisma_alani['Ortalama Görüntülenme Adedi'].fillna(0, inplace=True)

df_calisma_alani["Görüntülenmenin Satışa Dönüş Oranı"] = "0"

non_zero_mask = df_calisma_alani["Ortalama Görüntülenme Adedi"] != 0
df_calisma_alani.loc[non_zero_mask, "Görüntülenmenin Satışa Dönüş Oranı"] = round(
    (df_calisma_alani["MorhipoKodu"] / df_calisma_alani["Ortalama Görüntülenme Adedi"]) * 100, 2
)

# Tekrar kaydediyoruz
df_calisma_alani.to_excel("Nirvana.xlsx", index=False)





















import openpyxl
import gc
import re
from copy import copy

def duzenleme_islemleri(dosya_adi="Nirvana.xlsx", sayfa_adi="Sheet1"):
    # 1) Excel'i aç
    workbook = openpyxl.load_workbook(dosya_adi)
    if sayfa_adi not in workbook.sheetnames:
        print(f"'{sayfa_adi}' isminde bir sayfa bulunamadı.")
        return
    sheet = workbook[sayfa_adi]

    # İlk satırdan başlık -> sütun indekslerini elde et
    basliklar = {}
    for cell in sheet[1]:
        if cell.value:  # None değilse
            basliklar[cell.value] = cell.column

    # ------------------------------------------------
    # 1) "VaryasyonGittiGidiyorKodu" sütununu en sağdaki ilk boş sütuna kopyala,
    #    yeni sütuna "Kar Yüzdesi" adını ver + formül uygula
    # ------------------------------------------------
    vgk_baslik = "VaryasyonGittiGidiyorKodu"
    if vgk_baslik in basliklar:
        # "VaryasyonGittiGidiyorKodu" kolonunu bul
        stok_adedi_kolon = basliklar[vgk_baslik]
        # Yeni sütun yeri: En sağdaki ilk boş sütun
        yeni_sutun_index = sheet.max_column + 1

        # Eski hücre -> yeni hücre kopyalama (değer + stil)
        for row in range(1, sheet.max_row + 1):
            eski_hucre = sheet.cell(row=row, column=stok_adedi_kolon)
            yeni_hucre = sheet.cell(row=row, column=yeni_sutun_index)
            yeni_hucre.value = eski_hucre.value
            if eski_hucre.has_style:
                yeni_hucre._style = copy(eski_hucre._style)

        # Yeni sütuna başlık ismi
        sheet.cell(row=1, column=yeni_sutun_index).value = "Kar Yüzdesi"

        # Başlıkları tekrar güncelle (yeni sütun eklendi)
        basliklar = {}
        for cell in sheet[1]:
            if cell.value:
                basliklar[cell.value] = cell.column

        # "Satış Fiyatı" ve "Alış Fiyatı" sütun indekslerini al
        satis_fiyati_kolon = basliklar.get("SatisFiyati")
        alis_fiyati_kolon = basliklar.get("AlisFiyati")
        kar_yuzdesi_kolon = basliklar.get("Kar Yüzdesi")

        if satis_fiyati_kolon and alis_fiyati_kolon and kar_yuzdesi_kolon:
            for row in range(2, sheet.max_row + 1):
                sf = sheet.cell(row=row, column=satis_fiyati_kolon).value
                af = sheet.cell(row=row, column=alis_fiyati_kolon).value
                kar_hucre = sheet.cell(row=row, column=kar_yuzdesi_kolon)

                if sf and af:
                    try:
                        kar_orani = (sf - af) / sf
                        kar_hucre.value = kar_orani
                        kar_hucre.number_format = "0.00%"
                    except ZeroDivisionError:
                        kar_hucre.value = None

    # ------------------------------------------------
    # "Üründe Hareket Var mı?" adında yeni bir kolon oluşturma
    # ------------------------------------------------

    # ------------------------------------------------
    # 2) "Üründe Hareket Var mı?" adında yeni bir kolon oluştur
    # ------------------------------------------------
    yeni_kolon_sira = sheet.max_column + 1
    sheet.cell(row=1, column=yeni_kolon_sira).value = "Üründe Hareket Var mı?"

    # Başlıkları güncelleyelim
    basliklar = {}
    for cell in sheet[1]:
        if cell.value:  # None değilse
            basliklar[cell.value] = cell.column

    # İlgili kolon indeksleri
    resim_kolon = basliklar.get("Resim")
    stok_hersey_kolon = basliklar.get("Stok Adedi Her Şey Dahil")
    gunluk_ort_satis_kolon = basliklar.get("MorhipoKodu")
    dunun_satis_kolon = basliklar.get("Adet")

    for row in range(2, sheet.max_row + 1):
        sonuc_hucre = sheet.cell(row=row, column=yeni_kolon_sira)
        
        # 1) RESİM KONTROLÜ
        resim_deger = sheet.cell(row=row, column=resim_kolon).value if resim_kolon else None
        if not resim_deger or resim_deger == 0:
            sonuc_hucre.value = "Resim Yok"
            continue

        # 2) İLGİLİ KOLONLARI AL
        stok_deger = sheet.cell(row=row, column=stok_hersey_kolon).value if stok_hersey_kolon else 0
        gunluk_ort_satis = sheet.cell(row=row, column=gunluk_ort_satis_kolon).value if gunluk_ort_satis_kolon else 0
        dunun_satis = sheet.cell(row=row, column=dunun_satis_kolon).value if dunun_satis_kolon else 0
        
        # None gelirse 0 kabul edelim
        stok_deger = stok_deger if stok_deger else 0
        gunluk_ort_satis = gunluk_ort_satis if gunluk_ort_satis else 0
        dunun_satis = dunun_satis if dunun_satis else 0
        
        # 3) EVET KONTROLÜ
        # "Evet" => Günlük Ortalama Satış Adedi > 0 veya Dünün Satış Adedi > 0
        if (gunluk_ort_satis > 0) or (dunun_satis > 0) or (stok_deger > 0):
            sonuc_hucre.value = "Evet"
        
        # 4) HAYIR KONTROLÜ
        # "Hayır" => Stok <= 0 VE Günlük Ortalama Satış Adedi <= 0 VE Dünün Satış Adedi <= 0
        elif (stok_deger <= 0) and (gunluk_ort_satis <= 0) and (dunun_satis <= 0):
            sonuc_hucre.value = "Hayır"
        
        # 5) EĞER HİÇBİRİ TUTMADIYSA
        else:
            sonuc_hucre.value = "Hayır"
    # ------------------------------------------------
    # 3) TrendyolKodu ve VaryasyonTrendyolKodu kolonlarındaki
    #    veriler için ilk boşluktan sonraki kısmı temizle
    # ------------------------------------------------
    trendyol_kodu_baslik = "TrendyolKodu"
    varyant_trendyol_kodu_baslik = "VaryasyonTrendyolKodu"

    def remove_after_first_space(text):
        """
        "ABC DEF GHI" -> "ABC"
        İçerik boş veya None ise dokunma.
        """
        if not text or not isinstance(text, str):
            return text
        # En fazla 1 kez bölelim
        parts = text.split(" ", 1)
        return parts[0] if len(parts) > 1 else parts[0]

    for col_name in [trendyol_kodu_baslik, varyant_trendyol_kodu_baslik]:
        if col_name in basliklar:
            c_index = basliklar[col_name]
            for row in range(2, sheet.max_row + 1):
                hucre = sheet.cell(row=row, column=c_index)
                hucre.value = remove_after_first_space(hucre.value)

    # ------------------------------------------------
    # 4) İstenilen kolonları silelim + Kolon isimlerini değiştirelim
    # ------------------------------------------------

    # 4.2) "ToplamFiyat" kolonunu silelim (eğer varsa)
    if "ToplamFiyat" in basliklar:
        col_idx = basliklar["ToplamFiyat"]
        sheet.delete_cols(col_idx, 1)
        basliklar = {}
        for cell in sheet[1]:
            if cell.value:
                basliklar[cell.value] = cell.column

    # 4.3) Kolonları tekrar güncelleyip yeni isimler atayalım
    rename_map = {
        "UrunAdi": "Ürün Adı",
        "AlisFiyati": "Alış Fiyatı",
        "SatisFiyati": "Satış Fiyatı",
        "AramaTerimleri": "Resim Yüklenme Tarihi",
        "MorhipoKodu": "Günlük Ortalama Satış Adedi",
        "VaryasyonMorhipoKodu": "Depodaki Adetler",
        "N11Kodu": "Mevsim",
        "VaryasyonGittiGidiyorKodu": "Net Satış Tarihi ve Adedi",
        "TrendyolKodu": "Son Transfer Tarihi",
        "VaryasyonTrendyolKodu": "Son İndirim Tarihi",
        "StokAdedi": "Instagram Stok Adedi",
        "Adet": "Dünün Satış Adedi",
        "ListeFiyatı": "Liste Fiyatı",
    }

    for eski, yeni in rename_map.items():
        if eski in basliklar:
            col_idx = basliklar[eski]
            sheet.cell(row=1, column=col_idx).value = yeni

    # Sütunları yeniden oku
    basliklar = {}
    for cell in sheet[1]:
        if cell.value:
            basliklar[cell.value] = cell.column

    # ------------------------------------------------
    # 5) Son olarak, sütunların sırasını yeniden düzenle
    # ------------------------------------------------
    final_order = [

        "Ürün Adı",
        "Üründe Hareket Var mı?",
        "Instagram Stok Adedi",
        "Stok Adedi Her Şey Dahil",
        "Stok Adedi Site ve Vega",
        "Günlük Ortalama Satış Adedi",  
        "Dünün Satış Adedi",           
        "Ortalama Görüntülenme Adedi", 
        "Görüntülenmenin Satışa Dönüş Oranı",
        "Kaç Güne Biter Her Şey Dahil",
        "Kaç Güne Biter Site ve Vega",
        "Resim",
        "Alış Fiyatı",
        "Son Satın Alma Fiyatı",
        "Satış Fiyatı",
        "Liste Fiyatı",
        "Kar Yüzdesi",
        "Resim Yüklenme Tarihi",
        "Kategori",
        "GMT Stok Adedi",
        "SİTA Stok Adedi",
        "Mevsim",
        "Net Satış Tarihi ve Adedi",
        "Son Transfer Tarihi",
        "Son İndirim Tarihi",
        "Marka"
    ]

    # Mevcut tüm veriyi memory'e alıyoruz (list of dict)
    all_data = []
    headers_in_sheet = [cell.value for cell in sheet[1] if cell.value]

    for r in range(2, sheet.max_row + 1):
        row_dict = {}
        for header in headers_in_sheet:
            col_idx = basliklar.get(header)
            if col_idx:
                row_dict[header] = sheet.cell(row=r, column=col_idx).value
            else:
                row_dict[header] = None
        all_data.append(row_dict)

    # Sayfayı temizle (başlık satırı da dahil)
    sheet.delete_rows(1, sheet.max_row)

    # Yeni başlık satırını final_order'a göre yaz
    for col_idx, header_name in enumerate(final_order, start=1):
        sheet.cell(row=1, column=col_idx).value = header_name

    # all_data'daki her satır için final_order sırasına göre yaz
    for row_idx, row_data in enumerate(all_data, start=2):
        for col_idx, header_name in enumerate(final_order, start=1):
            value = row_data.get(header_name, None)
            sheet.cell(row=row_idx, column=col_idx).value = value

    # ------------------------------------------------------------
    # 6) En son: "Resim" kolonunda .jpeg kısaltma + "Ürün Adı" hücresine link ekle
    # ------------------------------------------------------------
    # Çünkü üstteki reorder işlemi hyperlink bilgisini siliyor.
    # Bu yüzden hyperlink atamayı en sona koyuyoruz.

    # Başlıkları tekrar toplayalım, zira rename oldu ("UrunAdi" -> "Ürün Adı" vb.)
    basliklar_son = {}
    for cell in sheet[1]:
        if cell.value:
            basliklar_son[cell.value] = cell.column

    # "Resim" -> "Ürün Adı"
    resim_header_son = "Resim"
    urun_adi_header_son = "Ürün Adı"

    if resim_header_son in basliklar_son and urun_adi_header_son in basliklar_son:
        resim_col_idx = basliklar_son[resim_header_son]
        urun_adi_col_idx = basliklar_son[urun_adi_header_son]

        for row in range(2, sheet.max_row + 1):
            resim_deger = sheet.cell(row=row, column=resim_col_idx).value
            if isinstance(resim_deger, str) and resim_deger.strip():
                # ";" varsa parçala, ilk linki al
                link_listesi = [lnk.strip() for lnk in resim_deger.split(';') if lnk.strip()]
                if link_listesi:
                    first_link = link_listesi[0]
                    # .jpeg sonrası silinsin
                    idx = first_link.lower().find(".jpeg")
                    truncated_value = first_link
                    if idx != -1:
                        idx += len(".jpeg")
                        truncated_value = first_link[:idx]

                    # Resim hücresine kısaltılmış değeri yaz
                    sheet.cell(row=row, column=resim_col_idx).value = truncated_value

                    # Kısaltılmış link http ile başlıyorsa "Ürün Adı" hücresine hyperlink
                    if truncated_value.lower().startswith("http"):
                        urun_adi_hucre = sheet.cell(row=row, column=urun_adi_col_idx)
                        # Aynı hücredeki metni koruyor, sadece link ekliyoruz.
                        urun_adi_hucre.hyperlink = truncated_value
                        # Not: Stil vermek istemiyorsak eklemeden geçiyoruz.

    # ------------------------------------------------
    # 7) Kaydet ve hafızayı temizle
    # ------------------------------------------------
    workbook.save(dosya_adi)
    gc.collect()



# Fonksiyonu çağır
if __name__ == "__main__":
    duzenleme_islemleri()
















import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

def tablo_duzenleme(dosya_adi="Nirvana.xlsx", sayfa_adi="Sheet1"):
    # 1) Excel'i aç
    wb = openpyxl.load_workbook(dosya_adi)
    
    if sayfa_adi not in wb.sheetnames:
        print(f"'{sayfa_adi}' isminde bir sayfa bulunamadı.")
        return
    
    sheet = wb[sayfa_adi]

    # ------------------------------------------------
    # 1) "UrunKodu" kolonunu hariç tüm verileri yatayda ortala
    # ------------------------------------------------
    basliklar = {}
    for cell in sheet[1]:
        if cell.value:
            basliklar[cell.value] = cell.column

    urun_kodu_kolon_index = basliklar.get("UrunKodu")

    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row):
        for cell_obj in row:
            if cell_obj.column != urun_kodu_kolon_index:
                cell_obj.alignment = Alignment(horizontal='center')

    # ------------------------------------------------
    # 2) Başlık satırının yüksekliğini 40 px yap
    # ------------------------------------------------
    sheet.row_dimensions[1].height = 40

    # ------------------------------------------------
    # 3) Başlık satırındaki verileri kalın, metni kaydırmalı,
    #    hem yatay hem dikey ortalı yapalım
    # ------------------------------------------------
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # ------------------------------------------------
    # 4) Freeze Panes: İlk satırı dondur
    # ------------------------------------------------
    sheet.freeze_panes = "A2"

    # ------------------------------------------------
    # 5) Sütun genişliklerini ayarla:
    #    - En az 120 px (yaklaşık 17.14 karakter genişliği)
    #    - Kolondaki en geniş veriye göre büyüt (karakter sayısı)
    # ------------------------------------------------
    min_width_chars = 120 / 7  # ~7 piksel/karakter
    for col in sheet.iter_cols(min_row=1, max_col=sheet.max_column, max_row=sheet.max_row):
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        
        # Başlık satırını (col[0]) atlayarak veri uzunluğunu ölç
        for cell_obj in col[1:]:
            value = cell_obj.value
            if value is not None:
                length = len(str(value))
                if length > max_length:
                    max_length = length

        # Sütun genişliğini belirle
        optimal_width = max(min_width_chars, max_length + 2)
        sheet.column_dimensions[col_letter].width = optimal_width

    # ------------------------------------------------
    # 6) "Marka" kolonunda "Sigara Ürün" içeren satırları Açık Yeşil yap
    # ------------------------------------------------
    marka_col_index = basliklar.get("Marka")
    if marka_col_index:
        max_col = sheet.max_column
        yesil_fill = PatternFill(start_color="A9D08E", end_color="A9D08E", fill_type="solid")

        for row_idx in range(2, sheet.max_row + 1):
            cell_value = sheet.cell(row=row_idx, column=marka_col_index).value
            if isinstance(cell_value, str) and "Sigara Ürün" in cell_value:
                for col_idx in range(1, max_col + 1):
                    sheet.cell(row=row_idx, column=col_idx).fill = yesil_fill

    # ------------------------------------------------
    # 7) Tabloyu "Beyaz Tablo Stili Açık 1" ile stillendirelim
    #    (bantlı satırlar için showRowStripes=True)
    # ------------------------------------------------
    max_row = sheet.max_row
    max_col = sheet.max_column
    start_col = get_column_letter(1)
    end_col = get_column_letter(max_col)
    tablo_araligi = f"{start_col}1:{end_col}{max_row}"

    my_table = Table(displayName="Tablom", ref=tablo_araligi)

    style = TableStyleInfo(
        name="TableStyleLight1", 
        showRowStripes=True,   # Şeritli (bantlı) satırlar
        showColumnStripes=False
    )
    my_table.tableStyleInfo = style
    sheet.add_table(my_table)

    # ------------------------------------------------
    # 8) Tüm kenarlıklar (thin) ekle
    #    Tablo aralığındaki her hücreyi ince siyah çerçeve yapıyoruz
    # ------------------------------------------------
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    for row_cells in sheet.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell_obj in row_cells:
            cell_obj.border = thin_border

    # ------------------------------------------------
    # 9) Değişiklikleri kaydet
    # ------------------------------------------------
    wb.save(dosya_adi)


if __name__ == "__main__":
    tablo_duzenleme()



import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from copy import copy

def hide_columns_by_header(sheet, headers_to_hide):
    """
    Verilen sheet'te, birinci satır başlığı headers_to_hide listesinde
    varsa o kolonu gizler.
    """
    max_col = sheet.max_column
    for col_index in range(1, max_col + 1):
        header_value = sheet.cell(row=1, column=col_index).value
        if header_value in headers_to_hide:
            col_letter = get_column_letter(col_index)
            sheet.column_dimensions[col_letter].hidden = True

def copy_sheet_style_structure(source_sheet, target_sheet):
    """
    Kaynaktaki sütun genişlikleri, satır yükseklikleri ve
    birleştirilmiş hücreleri (merges) hedef sayfaya kopyalar.
    """
    # 1) Sütun genişliklerini kopyala
    for col_letter, col_dim in source_sheet.column_dimensions.items():
        target_sheet.column_dimensions[col_letter].width = col_dim.width

    # 2) Satır yüksekliklerini kopyala
    for row_idx, row_dim in source_sheet.row_dimensions.items():
        target_sheet.row_dimensions[row_idx].height = row_dim.height

    # 3) Merge (birleştirme) aralıklarını kopyala
    for merged_range in source_sheet.merged_cells.ranges:
        target_sheet.merge_cells(str(merged_range))

def copy_row_with_style(source_sheet, target_sheet, src_row_idx, tgt_row_idx, max_col):
    """
    source_sheet'in src_row_idx satırını,
    target_sheet'in tgt_row_idx satırına, hücre stilleriyle birlikte kopyalar.
    """
    for c in range(1, max_col + 1):
        source_cell = source_sheet.cell(row=src_row_idx, column=c)
        target_cell = target_sheet.cell(row=tgt_row_idx, column=c)

        # Hücre değerini kopyala
        target_cell.value = source_cell.value

        # Stil kopyası (font, fill, border, alignment, number_format vb.)
        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.number_format = copy(source_cell.number_format)
            target_cell.protection = copy(source_cell.protection)
            target_cell.alignment = copy(source_cell.alignment)

def freeze_header(sheet):
    """
    Sayfanın ilk satırını (yani başlık satırını) dondurmak için 
    A2 hücresini 'freeze_panes' olarak ayarlar.
    """
    sheet.freeze_panes = "A2"

def apply_table_format(sheet, table_name="MyTable", style_name="TableStyleMedium9"):
    """
    Sayfadaki (A1'den başlayarak) mevcut veri aralığına 
    bir "Excel Tablosu" (Format as Table) ekler.
    Stil olarak 'TableStyleMedium9' kullanır 
    (veya isterseniz başka bir stil adını verebilirsiniz).
    
    - table_name: Tablonun workbook içindeki benzersiz adı.
    - style_name: "TableStyleMediumX", "TableStyleLightX" vs.
    """
    max_row = sheet.max_row
    max_col = sheet.max_column

    if max_row < 1 or max_col < 1:
        # Boş sayfa ise tablo oluşturulamaz
        return

    first_cell = "A1"
    last_cell = f"{get_column_letter(max_col)}{max_row}"
    table_ref = f"{first_cell}:{last_cell}"

    table = Table(displayName=table_name, ref=table_ref)
    style = TableStyleInfo(
        name=style_name,
        showRowStripes=True,   # Satırlar çizgili olsun
        showColumnStripes=False
    )
    table.tableStyleInfo = style

    sheet.add_table(table)

def filter_indirim_sheet_with_styles(
    source_sheet,
    target_sheet,
    instagram_stok_header="Instagram Stok Adedi",
    resim_header="Resim",
    urun_adi_header="Ürün Adı"
):
    """
    Özel olarak 'İndirim Raporu' sayfasının filtrelenmiş kopyasını oluşturur:
    - 'Instagram Stok Adedi' > 0
    - 'Resim' kolonundaki hücre boş olmayan satırları geçir
    - Her satırı kopyalarken stil bilgilerini de koru.
    - Kopyalama sonunda, 'Ürün Adı' kolonuna 'Resim' kolonundaki değeri hyperlink olarak ekle.
    """
    max_row = source_sheet.max_row
    max_col = source_sheet.max_column

    # Header (ilk satır)
    header_values = [source_sheet.cell(row=1, column=c).value for c in range(1, max_col+1)]

    # İlgili kolon indekslerini bulalım (1-bazlı)
    try:
        instagram_col_idx = header_values.index(instagram_stok_header) + 1
    except ValueError:
        instagram_col_idx = None

    try:
        resim_col_idx = header_values.index(resim_header) + 1
    except ValueError:
        resim_col_idx = None

    try:
        urun_adi_col_idx = header_values.index(urun_adi_header) + 1
    except ValueError:
        urun_adi_col_idx = None

    # Sayfanın sütun/row/merge yapılarını kopyala
    copy_sheet_style_structure(source_sheet, target_sheet)

    target_row_idx = 1
    for r in range(1, max_row + 1):
        if r == 1:
            # Header'ı (birinci satırı) doğrudan kopyala
            copy_row_with_style(source_sheet, target_sheet, r, target_row_idx, max_col)
            target_row_idx += 1
        else:
            # Filtre kontrolü
            pass_filter = True

            # 1) Instagram Stok Adedi > 0 mı?
            if instagram_col_idx is not None:
                val_instagram = source_sheet.cell(row=r, column=instagram_col_idx).value
                try:
                    if float(val_instagram) <= 0:
                        pass_filter = False
                except:
                    pass_filter = False  # sayısal değilse de at

            # 2) Resim kolonunda değer boş mu?
            if resim_col_idx is not None and pass_filter:
                val_resim = source_sheet.cell(row=r, column=resim_col_idx).value
                # Boş (None veya "") ise atla
                if val_resim is None or str(val_resim).strip() == "":
                    pass_filter = False

            if pass_filter:
                # Satırı kopyala
                copy_row_with_style(source_sheet, target_sheet, r, target_row_idx, max_col)

                # Ek olarak, "Ürün Adı" hücresine hyperlink verelim
                if (resim_col_idx is not None) and (urun_adi_col_idx is not None):
                    link_value = source_sheet.cell(row=r, column=resim_col_idx).value
                    product_name_cell = target_sheet.cell(row=target_row_idx, column=urun_adi_col_idx)
                    if link_value:
                        product_name_cell.hyperlink = str(link_value)

                target_row_idx += 1

def main():
    workbook_path = "Nirvana.xlsx"
    wb = openpyxl.load_workbook(workbook_path)

    # 1) "Sheet1" -> "Genel Rapor"
    source_sheet = wb["Sheet1"]
    source_sheet.title = "Genel Rapor"

    # 2) İki kopya oluştur
    rpt_sheet = wb.copy_worksheet(wb["Genel Rapor"])
    rpt_sheet.title = "RPT Raporu"

    indirim_sheet = wb.copy_worksheet(wb["Genel Rapor"])
    indirim_sheet.title = "İndirim Raporu"

    # 3) Kolon gizlemeleri
    # a) Genel Rapor
    hide_columns_by_header(
        wb["Genel Rapor"],
        ["Resim", "Marka"]
    )

    # b) RPT Raporu
    rpt_hide_cols = [
        "Stok Adedi Site ve Vega",
        "Ortalama Görüntülenme Adedi",
        "Kaç Güne Biter Site ve Vega",
        "Satış Fiyatı",
        "Resim Yüklenme Tarihi",
        "Kategori",
        "GMT Stok Adedi",
        "SİTA Stok Adedi",
        "Mevsim",
        "Son Transfer Tarihi",
        "Son İndirim Tarihi",
        "Resim",
        "Marka",
        "Liste Fiyatı",
    ]
    hide_columns_by_header(wb["RPT Raporu"], rpt_hide_cols)

    # c) İndirim Raporu (ilk etapta 'Resim' ve 'Marka' kolonlarını gizle)
    hide_columns_by_header(
        wb["İndirim Raporu"],
        ["Resim", "Marka", "Üründe Hareket Var mı?", "Stok Adedi Site ve Vega", "Kaç Güne Biter Site ve Vega", "Liste Fiyatı", "GMT Stok Adedi", "SİTA Stok Adedi", "Mevsim", "Net Satış Tarihi ve Adedi"]
    )

    # 4) "İndirim Raporu" sayfasında filtreleme
    temp_sheet_name = "Temp_Indirim"
    if temp_sheet_name in wb.sheetnames:
        del wb[temp_sheet_name]

    temp_sheet = wb.create_sheet(temp_sheet_name)

    filter_indirim_sheet_with_styles(
        source_sheet=wb["İndirim Raporu"],
        target_sheet=temp_sheet,
        instagram_stok_header="Instagram Stok Adedi",
        resim_header="Resim",
        urun_adi_header="Ürün Adı"
    )

    # Orijinal "İndirim Raporu" sayfasını sil, temp'i yeniden adlandır
    del wb["İndirim Raporu"]
    temp_sheet.title = "İndirim Raporu"

    # Tekrar 'Resim' ve 'Marka' kolonlarını gizle (yeni sayfa)
    hide_columns_by_header(
        wb["İndirim Raporu"],
        ["Resim", "Marka", "Üründe Hareket Var mı?", "Stok Adedi Site ve Vega", "Kaç Güne Biter Site ve Vega", "Liste Fiyatı", "GMT Stok Adedi", "SİTA Stok Adedi", "Mevsim", "Net Satış Tarihi ve Adedi"]
    )

    # 6) Başlık satırlarını dondur ve zoom ayarını %90 yap
    for sheet_name in ["Genel Rapor", "RPT Raporu", "İndirim Raporu"]:
        sh = wb[sheet_name]
        freeze_header(sh)
        sh.sheet_view.zoomScale = 90

    # 7) RPT Raporu ve İndirim Raporu sayfalarına tablo stili ekle
    apply_table_format(wb["RPT Raporu"], table_name="RPTTable", style_name="TableStyleMedium9")
    apply_table_format(wb["İndirim Raporu"], table_name="IndirimTable", style_name="TableStyleMedium9")

    # 8) Kaydet
    wb.save(workbook_path)

    # 9) 3 sayfada da 'Kar Yüzdesi' kolonundaki verileri 100 ile çarp
    for sheet_name in ["Genel Rapor", "RPT Raporu", "İndirim Raporu"]:
        sheet = wb[sheet_name]
        # Önce "Kar Yüzdesi" kolonunu bul
        kar_col_idx = None
        for col_index in range(1, sheet.max_column + 1):
            if sheet.cell(row=1, column=col_index).value == "Kar Yüzdesi":
                kar_col_idx = col_index
                break

        if kar_col_idx:
            # 2. satırdan son satıra kadar değeri 100 ile çarp
            for row_index in range(2, sheet.max_row + 1):
                val = sheet.cell(row=row_index, column=kar_col_idx).value
                if isinstance(val, (int, float)):
                    sheet.cell(row=row_index, column=kar_col_idx).value = val * 100

    # -------------------------------------------------------------------------
    # İSTENEN EKLEME: 3 sayfada da "Liste Fiyatı" ve "Satış Fiyatı" farkı %5'i aşıyorsa
    # "Alış Fiyatı" hücresinin yazı rengini kırmızı yapalım
    # -------------------------------------------------------------------------
    for sheet_name in ["Genel Rapor", "RPT Raporu", "İndirim Raporu"]:
        sheet = wb[sheet_name]

        # Kolon indekslerini bulalım
        liste_fiyati_idx = None
        satis_fiyati_idx = None
        alis_fiyati_idx = None

        for col_index in range(1, sheet.max_column + 1):
            header_val = sheet.cell(row=1, column=col_index).value
            if header_val == "Liste Fiyatı":
                liste_fiyati_idx = col_index
            elif header_val == "Satış Fiyatı":
                satis_fiyati_idx = col_index
            elif header_val == "Alış Fiyatı":
                alis_fiyati_idx = col_index

        # Eğer bu kolonlar varsa kontrol edelim
        if liste_fiyati_idx and satis_fiyati_idx and alis_fiyati_idx:
            for row_index in range(2, sheet.max_row + 1):
                try:
                    list_val = float(sheet.cell(row=row_index, column=liste_fiyati_idx).value or 0)
                    satis_val = float(sheet.cell(row=row_index, column=satis_fiyati_idx).value or 0)
                    
                    # Liste Fiyatı 0 değilse hesaplama yapalım
                    if list_val != 0:
                        fark = list_val - satis_val
                        yuzde = (fark / list_val) * 100
                        if yuzde > 5:
                            # Alış Fiyatı hücresinin yazı rengini kırmızıya çevir
                            cell_alis = sheet.cell(row=row_index, column=alis_fiyati_idx)
                            existing_font = copy(cell_alis.font)
                            existing_font.color = "FF0000"  # Kırmızı
                            cell_alis.font = existing_font
                except:
                    # Herhangi bir hatada (None, string vs.), atla
                    pass

    # Tüm değişiklikleri kaydet
    wb.save(workbook_path)

if __name__ == "__main__":
    main()








import openpyxl
from openpyxl.utils import get_column_letter

# Dosyayı yükle
workbook = openpyxl.load_workbook("Nirvana.xlsx")
worksheet = workbook["RPT Raporu"]

# Başlık satırındaki kolon isimlerini ve indekslerini tespit edelim (1. satır)
kolon_indexleri = {}
for idx, hucre in enumerate(worksheet[1], start=1):
    kolon_indexleri[hucre.value] = idx

# "Liste Fiyatı" ve "Satış Fiyatı" kolonlarının indekslerini alalım
liste_fiyati_col = kolon_indexleri.get("Liste Fiyatı")
satis_fiyati_col = kolon_indexleri.get("Satış Fiyatı")

if liste_fiyati_col is None or satis_fiyati_col is None:
    raise ValueError("Gerekli kolonlardan biri bulunamadı: 'Liste Fiyatı' veya 'Satış Fiyatı'.")

# Yeni kolonun (İndirim Oranı) ekleneceği indeks
yeni_kolon = worksheet.max_column + 1
worksheet.cell(row=1, column=yeni_kolon, value="İndirim Oranı")

# Veriler üzerinde işlem yapalım (başlık satırından sonraki satırlar)
for satir in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
    liste_deger = satir[liste_fiyati_col - 1].value  # 0 tabanlı indeksleme
    satis_deger = satir[satis_fiyati_col - 1].value

    # Liste Fiyatı değerinin geçerli ve sıfırdan farklı olduğundan emin olalım
    if liste_deger is not None and satis_deger is not None and liste_deger != 0:
        indirim_orani = (liste_deger - satis_deger) / liste_deger * 100
        worksheet.cell(row=satir[0].row, column=yeni_kolon, value=indirim_orani)
        
        # Eğer indirim oranı 10 veya daha büyükse satırı gizle
        if indirim_orani >= 10:
            worksheet.row_dimensions[satir[0].row].hidden = True
    else:
        # Geçerli veri yoksa yeni kolonda boş bırak
        worksheet.cell(row=satir[0].row, column=yeni_kolon, value=None)

# Yeni oluşturduğumuz "İndirim Oranı" kolonunu gizleyelim
col_letter = get_column_letter(yeni_kolon)
worksheet.column_dimensions[col_letter].hidden = True

# Sonucu yeni bir dosyaya kaydedelim
workbook.save("Nirvana.xlsx")









import openpyxl
from openpyxl.comments import Comment

# Açıklama metinlerini içeren sözlük
aciklamalar = {
    "Ürün Adı": "Ürünün satıştaki adını belirtir.",
    "Üründe Hareket Var mı?": 'Bu alanda "Hayır" seçeneğin işareti kaldırılırsa, üründen hiçbir yerde yok ise ve aynı zamanda satışında da yakın zamanda bir hareket yok ise bu satırları gizle anlamına gelecektir.',
    "Instagram Stok Adedi": "Satıştaki stok adedini belirtir.",
    "Stok Adedi Her Şey Dahil": "Satıştaki stok, depodaki stok, çuvaldaki stok, siparişteki stokların toplamını belirtir.",
    "Stok Adedi Site ve Vega": "Satıştaki stokla depoda satışa hazır ürünlerin stoğunu belirtir.",
    "Günlük Ortalama Satış Adedi": "Son 7 günde ürünün satışta kaldığı güne göre ortalama bir satış adedi belirtir.",
    "Dünün Satış Adedi": "Ürünün dün kaç adet sattığını belirtir.",
    "Ortalama Görüntülenme Adedi": "Son 7 günde ürünün satışta kaldığı güne göre ortalama bir görüntülenme adedi belirtir.",
    "Görüntülenmenin Satışa Dönüş Oranı": "Son 7 günde ürünün satışta kaldığı güne göre ortalama bir görüntülenme adedi belirlenir ve bu görüntülenme adedine göre yine son 7 gündeki ortalama satışı baz alınarak bir yüzde hesabı çıkarılır.",
    "Kaç Güne Biter Her Şey Dahil": "Her dahil stokların, ortalama satış adedi baz alınarak kaç güne biteceğin belirtir.",
    "Kaç Güne Biter Site ve Vega": "Sadece satıştaki ve depoda satışa hazır stokların ortalama satış adedi baz alınarak kaç güne biteceğini belirtir.",
    "Alış Fiyatı": "Ürünün siteye geçirilen alış fiyatını belirtir. NOT : Ürünü alırkenki alış fiyatını yansıtmayabilir.",
    "Son Satın Alma Fiyatı": "Ürünün sipariş verirken yazılan son alış fiyatını belirtir.",
    "Satış Fiyatı": "Ürünün aktif satış fiyatını belirtir.",
    "Liste Fiyatı": "Ürünün alış fiyatına oranla olması gereken fiyatını belirtir.",
    "Kar Yüzdesi": "Ürünün yüzde kaç karla sattığını belirtir ve hesaplaması şu şekildedir (Satış - Alış) / Satış) * 100",
    "Resim Yüklenme Tarihi": "Ürünün resminin yüklendiği bir nevi satışa açıldığı tarihi belirtir.",
    "Kategori": "Ürünün ana kategorisini belirtir.",
    "GMT Stok Adedi": "Üründen çuvalda kaç adet olduğunu belirtir.",
    "SİTA Stok Adedi": "Ürünün siparişte kaç adet olduğunu belirtir.",
    "Mevsim": "Ürünün mevsimini belirtir.",
    "Net Satış Tarihi ve Adedi": "Ürünün tüm renk ve bedenlerinin aynı anda satışta olduğu son günün satış tarihi ve adedini belirtir.",
    "Son Transfer Tarihi": "Ürünün Instagram depoya en son ne zaman transfer edildiğini belirtir.",
    "Son İndirim Tarihi": "Ürüne en son ne zaman indirim yapıldığını belirtir."
}

# Dosyayı aç
workbook = openpyxl.load_workbook("Nirvana.xlsx")

# Tüm sayfaları tarayalım
for worksheet in workbook.worksheets:
    # İlk satırdaki hücrelerde başlıklar olduğunu varsayıyoruz
    for cell in worksheet[1]:
        if cell.value in aciklamalar:
            # Yorum objesini oluşturup genişlik ve yükseklik ayarını yapalım
            comment = Comment(text=aciklamalar[cell.value], author="Auto")
            comment.width = 200   # 200 piksel genişlik
            comment.height = 200  # 200 piksel yükseklik
            cell.comment = comment

# Sonuç dosyasını kaydedelim
workbook.save("Nirvana.xlsx")
