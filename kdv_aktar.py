# -*- coding: utf-8 -*-
"""
BenimPOS dışa aktarım dosyasındaki KDV oranlarını ve Fiyatları,
HARUN CALISMA dosyasındaki ilgili sütunlara barkod eşleştirmesiyle yazar.
"""

import re
import shutil
import pandas as pd
from openpyxl import load_workbook

# ---------------- AYARLAR ----------------
POS_DOSYA     = "Ürünler_Dışa_Aktar_-__BenimPOS.xlsx"
POS_SAYFA     = "Sheet1"
POS_BASLIK    = 1          # başlıklar 2. satırda (0'dan sayınca 1)
POS_BARKOD    = "Ürün barkodu"
POS_KDV       = "KDV"
POS_FIYAT     = "Fiyat 1"  # Fiyatın çekileceği sütun ("Alış Fiyatı" vb. de yapabilirsin)

HEDEF_DOSYA   = "HARUN_CALISMA__1_.xlsx"
HEDEF_SAYFA   = "Ürünlerinizi Burada Listeleyin"
HEDEF_BARKOD  = "Barkod"
HEDEF_KDV     = "KDV Oranı"
HEDEF_FIYAT   = "Fiyat"    # KDV'nin yanına eklenecek yeni başlık

CIKTI_DOSYA   = "HARUN_CALISMA_KDV_FIYAT_GUNCEL.xlsx"
# -----------------------------------------

def normalize(barkod):
    """Barkodu karşılaştırılabilir hale getirir:
    tırnak/boşluk/harf temizler, baştaki sıfırları atar."""
    rakamlar = re.sub(r"\D", "", str(barkod))
    return rakamlar.lstrip("0")

# 1) POS dosyasından barkod -> (KDV, Fiyat) sözlüğü kur
pos = pd.read_excel(POS_DOSYA, sheet_name=POS_SAYFA, header=POS_BASLIK, dtype=str)
veri_sozlugu = {}
for barkod, kdv, fiyat in zip(pos[POS_BARKOD], pos[POS_KDV], pos[POS_FIYAT]):
    anahtar = normalize(barkod)
    # Eğer kdv veya fiyat değerinden en az biri varsa sözlüğe al
    if anahtar and (pd.notna(kdv) or pd.notna(fiyat)) and anahtar not in veri_sozlugu:
        veri_sozlugu[anahtar] = {"kdv": kdv, "fiyat": fiyat}
print(f"POS dosyasından {len(veri_sozlugu)} benzersiz barkod okundu.")

# 2) Hedef dosyanın kopyasını aç (orijinal bozulmasın)
shutil.copy(HEDEF_DOSYA, CIKTI_DOSYA)
wb = load_workbook(CIKTI_DOSYA)
ws = wb[HEDEF_SAYFA]

# Sütun numaralarını başlık satırından bul
basliklar = {str(c.value).strip(): c.column for c in ws[1] if c.value}
col_barkod = basliklar[HEDEF_BARKOD]
col_kdv = basliklar[HEDEF_KDV]
col_fiyat = col_kdv + 1  # Fiyatı KDV'nin hemen sağına ekleyeceğiz

# Başlığı yazalım
ws.cell(row=1, column=col_fiyat).value = HEDEF_FIYAT

# 3) Satır satır eşleştir ve yaz
bulunan, bulunamayan = 0, []
for satir in range(2, ws.max_row + 1):
    hucre_barkod = ws.cell(row=satir, column=col_barkod).value
    if hucre_barkod is None:
        continue

    # Bir hücrede boşlukla ayrılmış birden fazla barkod olabilir
    yeni_kdv, yeni_fiyat = None, None
    for parca in str(hucre_barkod).split():
        anahtar = normalize(parca)
        if anahtar in veri_sozlugu:
            yeni_kdv = veri_sozlugu[anahtar]["kdv"]
            yeni_fiyat = veri_sozlugu[anahtar]["fiyat"]
            break

    if yeni_kdv is not None or yeni_fiyat is not None:
        # KDV Yazımı
        if pd.notna(yeni_kdv):
            try:
                yeni_kdv = float(yeni_kdv)
                if yeni_kdv == int(yeni_kdv):
                    yeni_kdv = int(yeni_kdv)
            except (TypeError, ValueError):
                pass
            ws.cell(row=satir, column=col_kdv).value = yeni_kdv
            
        # Fiyat Yazımı
        if pd.notna(yeni_fiyat):
            try:
                # Fiyatlar ondalıklı olabildiği için float'a çeviriyoruz
                yeni_fiyat = float(yeni_fiyat)
            except (TypeError, ValueError):
                pass
            ws.cell(row=satir, column=col_fiyat).value = yeni_fiyat

        bulunan += 1
    else:
        bulunamayan.append((satir, str(hucre_barkod).strip()))

wb.save(CIKTI_DOSYA)

# 4) Rapor
print(f"Güncellenen satır : {bulunan}")
print(f"Eşleşmeyen satır  : {len(bulunamayan)}")
if bulunamayan:
    rapor = pd.DataFrame(bulunamayan, columns=["Excel Satırı", "Barkod"])
    rapor.to_excel("eslesmeyen_barkodlar.xlsx", index=False)
    print("Eşleşmeyenler -> eslesmeyen_barkodlar.xlsx")
print(f"Çıktı dosyası     : {CIKTI_DOSYA}")