import pandas as pd

# Giriş dosyası adı
input_file = "finansallar.xlsx"
# Çıktı dosyası adı
output_file = "likidite_analizleri.xlsx"

# Hesaplamalar için kullanılacak satır isimleri (itemDescTr sütunundaki değerler)
row_names = {
    "donen_varliklar": "Dönen Varlıklar",
    "stoklar": "Stoklar",
    "nakit": "Nakit ve Nakit Benzerleri",
    "kisavade_yukumlulukler": "Kısa Vadeli Yükümlülükler"
}

# Excel dosyasındaki tüm sheet isimlerini alıyoruz.
xls = pd.ExcelFile(input_file)
sheet_names = xls.sheet_names

# Sonuçları saklamak için boş sözlükler; 
# her oran için DataFrame oluşturacağız: index = hisse (sheet adı), columns = çeyrek tarihleri.
cari_oran_dict = {}

nakit_oran_dict = {}

def ensure_series(val, columns):
    """
    Eğer 'val' scalar ise, verilen sütunları index olarak kullanarak değeri tekrarlayan pandas Series oluştur.
    Eğer zaten Series ise olduğu gibi döndür.
    """
    if isinstance(val, pd.Series):
        return val
    else:
        return pd.Series([val] * len(columns), index=columns)

def to_numeric_series(value, columns=None):
    """
    Gelen değeri önce Series'e çevirir (eğer değilse) ve ardından numerik formata dönüştürür.
    Eğer columns parametresi sağlanırsa, Series'i bu index ile oluşturur.
    """
    if not isinstance(value, pd.Series):
        if columns is not None:
            value = pd.Series([value] * len(columns), index=columns)
        else:
            value = pd.Series([value])
    return pd.to_numeric(value, errors='coerce')

# Her bir sheet üzerinde döngü
for sheet in sheet_names:
    # Her sheet’i DataFrame olarak oku
    df = pd.read_excel(input_file, sheet_name=sheet)
    # itemDescTr sütununu index olarak kullanmak işimizi kolaylaştırır.
    df.set_index("itemDescTr", inplace=True)
    
    # Çeyrek tarihleri sütunlarını tespit et (örneğin '2021/3', '2021/6', ...)
    quarter_columns = [col for col in df.columns if isinstance(col, str) and "/" in col]
    
    # Eğer quarter_columns boşsa, bu sheet'i atla.
    if not quarter_columns:
        print(f"Sheet {sheet}: Çeyrek sütunları bulunamadı, atlanıyor.")
        continue

    try:
        donen_varliklar_raw = df.loc[row_names["donen_varliklar"], quarter_columns]
        stoklar_raw = df.loc[row_names["stoklar"], quarter_columns]
        nakit_raw = df.loc[row_names["nakit"], quarter_columns]
        kisavade_raw = df.loc[row_names["kisavade_yukumlulukler"], quarter_columns]
    except KeyError as e:
        print(f"Sheet {sheet}: {e} bulunamadı. Bu hisse için hesaplama atlanıyor.")
        continue

    # Önce değerleri Series haline getiriyoruz, ardından numerik değere çeviriyoruz.
    donen_varliklar = to_numeric_series(donen_varliklar_raw, quarter_columns)
    stoklar = to_numeric_series(stoklar_raw, quarter_columns)
    nakit = to_numeric_series(nakit_raw, quarter_columns)
    kisavade = to_numeric_series(kisavade_raw, quarter_columns)

    # Eğer Kısa Vadeli Yükümlülükler (kisavade) tüm hücrelerde boşsa, hesaplamaya devam etmiyoruz.
    if kisavade.isna().all():
        print(f"Sheet {sheet}: Kısa Vadeli Yükümlülükler tüm değerler boş, hesaplama atlanıyor.")
        continue

    # Eğer stoklar hücrelerinde boşluk varsa, bunları 0 kabul edelim (stok raporu verilmiyorsa)
    stoklar = stoklar.fillna(0)

    # Hesaplamalar
    cari_oran = donen_varliklar / kisavade
    asit_test_oran = (donen_varliklar - stoklar) / kisavade
    nakit_oran = nakit / kisavade

    # Sonuçları sözlüklere ekle. Her sheet bir hisse adı olarak kullanılacak.
    cari_oran_dict[sheet] = cari_oran
   
    nakit_oran_dict[sheet] = nakit_oran

# Her oran için DataFrame oluşturma:
cari_oran_df = pd.DataFrame(cari_oran_dict).T  # Satır indexi: hisse, sütunlar: çeyrekler

nakit_oran_df = pd.DataFrame(nakit_oran_dict).T

# Çeyrek sütunlarını kronolojik sıraya sokmak isteyebilirsiniz.
def sort_quarters(columns):
    # 'YYYY/Çeyrek' formatında sıralama yapıyoruz
    return sorted(columns, key=lambda x: (int(x.split("/")[0]), int(x.split("/")[1])))

sorted_columns = sort_quarters(cari_oran_df.columns.tolist())
cari_oran_df = cari_oran_df[sorted_columns]

nakit_oran_df = nakit_oran_df[sorted_columns]

# Sonuçları Excel’e yazma (her oran için ayrı bir sheet)
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    cari_oran_df.to_excel(writer, sheet_name="Cari Oran")
    
    nakit_oran_df.to_excel(writer, sheet_name="Nakit Oran")

print(f"Hesaplamalar tamamlandı. Sonuçlar {output_file} dosyasına kaydedildi.")
