# -*- coding: utf-8 -*-
import pandas as pd
import openpyxl
from openpyxl.descriptors.base import Set
import glob
import sys
import warnings
import re

# ===================================================================
# Monkeypatch openpyxl Set validation to bypass "invalid XML" errors
# caused by strict rule validation on ConditionalFormatting in the template
# ===================================================================
original_set = Set.__set__
def bypass_set(self, instance, value):
    try:
        original_set(self, instance, value)
    except ValueError:
        pass # Ignore validation errors
Set.__set__ = bypass_set
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ===================================================================
# SUTUN HARITALARI (SABIT)
# ===================================================================

# And / IOS sekmeleri icin harita
GENEL_COL_MAP = {
    ('Active users',    'Start'):    2,
    ('Active users',    'Complete'): 3,
    ('Active users',    'Fail'):     4,
    ('Active users',    'Tryagain'): 5,
    ('avarage_attempt', 'Complete'): 12,
    ('avg_thinktime',   'Complete'): 19,
    ('Total users',     'ADs'):      20, # Avg.Rewarded
    ('coin_owned',      'Complete'): 21, # Owned Coin
}

# Booster sekmeleri icin harita
BOOSTER_COL_MAP = {
    ('booster_a_used',  'Complete'): 3,
    ('booster_b_used',  'Complete'): 4,
    ('booster_c_used',  'Complete'): 5,
    ('booster_a_used',  'Fail'):     6,
    ('booster_b_used',  'Fail'):     7,
    ('booster_c_used',  'Fail'):     8,
    ('booster_a_owned', 'Complete'): 15,
    ('booster_b_owned', 'Complete'): 16,
    ('booster_c_owned', 'Complete'): 17,
    ('booster_a_owned', 'Fail'):     18,
    ('booster_b_owned', 'Fail'):     19,
    ('booster_c_owned', 'Fail'):     20,
    ('Total users',     'ADs'):      21, # Total Rewarded
}

def get_max_level_from_user():
    """Kullanicidan kacinci levele kadar islem yapilacagini sor."""
    # Test ortaminda input beklentisini karsilamak icin (mockable)
    while True:
        try:
            val = input("Kacinci levele kadar analiz yapilacak? (Ornek: 300) : ")
            return int(val.strip())
        except (ValueError, EOFError):
            return 300 # Varsayilan

def detect_funnel_type(df_raw):
    """Dosyanin meta basliginda funnel turunu tespit et."""
    for idx in range(min(10, len(df_raw))):
        row_str = " ".join(df_raw.iloc[idx].dropna().astype(str))
        for keyword in ['Complete', 'Fail', 'Tryagain', 'Start', 'ADs']:
            if keyword.lower() in row_str.lower():
                return keyword
    return 'Unknown'


def read_platform_data(f):
    """
    Excel dosyasini okur, Android ve iOS verilerini ayri DataFrame olarak dondurur.
    Doner: (funnel_type, df_android, df_ios)
    """
    df_raw = pd.read_excel(f, header=None)
    funnel_type = detect_funnel_type(df_raw)

    os_row_idx = None
    level_row_idx = None
    for idx in range(min(15, len(df_raw))):
        row = df_raw.iloc[idx]
        for cell in row:
            if str(cell).strip() == 'Operating system' and os_row_idx is None:
                os_row_idx = idx
            if str(cell).strip() == 'Level' and level_row_idx is None:
                level_row_idx = idx

    if os_row_idx is None or level_row_idx is None:
        print(f"   UYARI: '{f}' -> 'Operating system' veya 'Level' bulunamadi, atlaniyor.")
        return funnel_type, pd.DataFrame(), pd.DataFrame()

    os_row_vals  = df_raw.iloc[os_row_idx].fillna('').astype(str).str.strip().tolist()
    level_row_vals = [str(x).strip() for x in df_raw.iloc[level_row_idx].tolist()]

    if level_row_idx == os_row_idx:
        df_flat = df_raw.iloc[level_row_idx + 1:].copy()
        seen = {}
        unique_cols = []
        for c in level_row_vals:
            if c in seen:
                seen[c] += 1
                unique_cols.append(f"{c}_{seen[c]}")
            else:
                seen[c] = 0
                unique_cols.append(c)
        df_flat.columns = unique_cols
        df_flat = df_flat.dropna(subset=['Level'])

        df_and  = df_flat[df_flat['Operating system'].astype(str).str.contains('Android', case=False, na=False)].copy()
        df_ios  = df_flat[df_flat['Operating system'].astype(str).str.contains('iOS',     case=False, na=False)].copy()
        
        df_and  = df_and.drop(columns=['Operating system'], errors='ignore').reset_index(drop=True)
        df_ios  = df_ios.drop(columns=['Operating system'], errors='ignore').reset_index(drop=True)
        return funnel_type, df_and, df_ios

    else:
        level_col_idx = next((i for i, v in enumerate(level_row_vals) if v == 'Level'), None)
        if level_col_idx is None:
            print(f"   UYARI: '{f}' -> Level sutun indeksi bulunamadi.")
            return funnel_type, pd.DataFrame(), pd.DataFrame()

        and_indices = [i for i, v in enumerate(os_row_vals) if v == 'Android']
        ios_indices = [i for i, v in enumerate(os_row_vals) if v == 'iOS']
        data_start  = level_row_idx + 1

        def extract(col_indices):
            cols = [level_col_idx] + col_indices
            sub  = df_raw.iloc[data_start:, cols].copy()
            raw_names = [level_row_vals[level_col_idx]] + [level_row_vals[i] for i in col_indices]
            seen = {}
            unique_names = []
            for n in raw_names:
                if n in seen:
                    seen[n] += 1
                    unique_names.append(f"{n}_{seen[n]}")
                else:
                    seen[n] = 0
                    unique_names.append(n)
            sub.columns = unique_names
            return sub.dropna(subset=['Level']).reset_index(drop=True)

        return funnel_type, extract(and_indices), extract(ios_indices)


def safe_concat(df_list):
    """Farkli/duplikat sutun isimlerine sahip DataFrame listesini birlestir."""
    if not df_list: return pd.DataFrame()
    if len(df_list) == 1: return df_list[0]
    
    cleaned = []
    for df in df_list:
        if df.empty: continue
        df = df.reset_index(drop=True).copy()
        cols = list(df.columns)
        seen = {}
        unique_cols = []
        for c in cols:
            if c in seen:
                seen[c] += 1
                unique_cols.append(f"{c}_{seen[c]}")
            else:
                seen[c] = 0
                unique_cols.append(str(c))
        df.columns = unique_cols
        cleaned.append(df)
    
    if not cleaned: return pd.DataFrame()
    return pd.concat(cleaned, ignore_index=True, sort=False)


def format_and_fill_levels(df_sub, max_level):
    """Level bazinda topla, max_level'a kadar eksiksiz tablo dondur."""
    if df_sub.empty or 'Level' not in df_sub.columns:
        return pd.DataFrame({'Level': range(1, max_level + 1)})

    df_sub = df_sub.copy()
    df_sub['Level'] = pd.to_numeric(df_sub['Level'], errors='coerce')
    df_sub = df_sub.dropna(subset=['Level']).copy()
    df_sub['Level'] = df_sub['Level'].astype(int)
    df_sub = df_sub[df_sub['Level'] <= max_level]
    
    for col in df_sub.columns:
        if col != 'Level':
            df_sub[col] = pd.to_numeric(df_sub[col], errors='coerce').fillna(0)

    numeric_cols = df_sub.select_dtypes(include='number').columns.tolist()
    if 'Level' in numeric_cols:
        numeric_cols.remove('Level')

    grouped = df_sub.groupby('Level')[numeric_cols].sum().reset_index()
    grouped = grouped.sort_values(by='Level')

    all_levels = pd.DataFrame({'Level': range(1, max_level + 1)})
    merged = pd.merge(all_levels, grouped, on='Level', how='left')
    merged.fillna(0, inplace=True)
    return merged


def find_level_rows(ws):
    """Sablon sekmesinde Level -> satir_no haritasini olustur."""
    mapping = {}
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, (int, float)):
            mapping[int(v)] = r
    return mapping


def shift_formula(formula, shift):
    """
    Excel formulundeki satir referanslarini (N3 -> N4 gibi) kaydirir.
    Regex ile hucre referanslarini bulur ve satir numarasini artirir.
    """
    if not formula or not isinstance(formula, str) or not formula.startswith('='):
        return formula

    def replace_match(match):
        prefix = match.group(1) # Harf(ler) ve varsa $ isareti
        row_num = match.group(2) # Satir numarasi
        
        # Eger satir numarasi $ ile sabitlenmisse ($3 gibi), kaydirma
        if prefix.endswith('$'):
            return f"{prefix}{row_num}"
        
        new_row = int(row_num) + shift
        return f"{prefix}{new_row}"

    # Regex: Hucre referanslarini bulur. 
    # Grp 1: Sutun harfleri (opsiyonel $)
    # Grp 2: Satir numarasi
    # Negatif lookahead ile sonrasinda ( veya ! gelmediginden emin olmaya gerek yok 
    # cunku satir no her zaman harf grubundan sonra gelir.
    pattern = r'(\$?[A-Z]+\$?)([0-9]+)'
    return re.sub(pattern, replace_match, formula)


def write_tab_data(wb, sheet_name, data_by_funnel, col_map, max_level, revenue_df=None):
    """
    Belirli bir sekme icin verileri funnel haritasina gore yazdirir.
    Level 1'den max_level'a kadar tüm satırları doldurur.
    """
    if sheet_name not in wb.sheetnames:
        print(f"   ! Uyari: '{sheet_name}' sekmesi yok, atlaniyor.")
        return

    ws = wb[sheet_name]
    
    # 1. TEMIZLIK: Sadece veri yazilacak sutunlari temizle (Row 3-1000)
    cols_to_clear = set(col_map.values())
    if sheet_name in ['And', 'IOS']:
        cols_to_clear.add(14) # Total $
        # Sadece K sutununa (11. sutun) dokunulmasin dedigi icin onu temizlik listesinden cikaralim
        if 11 in cols_to_clear:
            cols_to_clear.remove(11)
    
    for r in range(3, 1001):
        for c in cols_to_clear:
            if sheet_name in ['And', 'IOS'] and c == 11:
                continue # K sutununa dokunma
            ws.cell(row=r, column=c).value = None

    level_rows = find_level_rows(ws)

    # 2. Tum huni tiplerini hazırla ve indexle
    all_funnel_types = ['Start', 'Complete', 'Fail', 'Tryagain', 'ADs']
    processed_dfs = {}
    for ftype in all_funnel_types:
        df = data_by_funnel.get(ftype, pd.DataFrame())
        df_filled = format_and_fill_levels(df, max_level)
        processed_dfs[ftype] = df_filled.set_index('Level')

    # 3. Gelir verisini hazırla
    if revenue_df is not None:
        rev_filled = format_and_fill_levels(revenue_df, max_level)
        if 'Total $' not in rev_filled.columns:
            rev_filled['Total $'] = 0
        revenue_df = rev_filled.set_index('Level')

    # 4. Yazdirma dongusu
    for level in range(1, max_level + 1):
        if level not in level_rows:
            target_row = level + 2
            ws.cell(row=target_row, column=1).value = level
            level_rows[level] = target_row
        
        target_row = level_rows[level]

        # Huni verilerini yaz
        for ftype, df in processed_dfs.items():
            if level in df.index:
                r_val = df.loc[level]
                for col_name in df.columns:
                    base_name = col_name.split('_')[0] if '_' in col_name and col_name.split('_')[-1].isdigit() else col_name
                    key = (base_name, ftype)
                    if key in col_map:
                        col_idx = col_map[key]
                        if sheet_name in ['And', 'IOS'] and col_idx == 11:
                            continue # K sutununa dokunma
                        val = r_val[col_name]
                        ws.cell(row=target_row, column=col_idx).value = val

        # Gelir verisini yaz (Sadece And/IOS sekmelerinde N sütunu)
        if revenue_df is not None:
            if level in revenue_df.index:
                ws.cell(row=target_row, column=14).value = revenue_df.loc[level].get('Total $', 0)

    # 5. FORMUL UZATMA: Row 3'teki formulleri max_level'a kadar kopyala ve SATIRLARI KAYDIR
    formula_cols = []
    for c in range(2, ws.max_column + 1):
        if sheet_name in ['And', 'IOS'] and c == 11:
            continue # K sutununa dokunma
            
        cell = ws.cell(row=3, column=c)
        if isinstance(cell.value, str) and cell.value.startswith('='):
            formula_cols.append((c, cell.value))
    
    if formula_cols:
        for level in range(2, max_level + 1): # Level 2 ve uzeri (Row 4+)
            t_row = level_rows.get(level, level + 2)
            shift = t_row - 3 # Row 3'ten baslayarak ne kadar asagi kaydigimiz
            for c_idx, f_str in formula_cols:
                shifted_f = shift_formula(f_str, shift)
                ws.cell(row=t_row, column=c_idx).value = shifted_f

    print(f"   OK '{sheet_name}' tablosu yazildi (Level 1-{max_level}).")


def main():
    print("=======================================================")
    print(" Hole Pool Veri Isleme Scripti Baslatiliyor...")
    print("=======================================================\n")

    max_level = get_max_level_from_user()
    print(f"\n[OK] Analiz {max_level}. levele kadar yapilacak.\n")

    # 1. Dosyalari bul (Iceriginde # hole-pool olan .xlsx dosyalarini tara)
    print("1. Veri dosyalari (xlsx) taraniyor...")
    all_xlsx = glob.glob("*.xlsx")
    files = []
    date_dict = {}

    for f in all_xlsx:
        if f.startswith("Copy of Hole_Pool") or f == "Hole_Pool_Otomatik_Analiz_v1.xlsx":
            continue
        try:
            df_check = pd.read_excel(f, nrows=10, header=None)
            content_str = " ".join(df_check.fillna('').astype(str).values.flatten())
            if "# hole-pool" in content_str.lower():
                files.append(f)
                found_date = None
                for idx, row in df_check.iterrows():
                    row_str = " ".join(row.dropna().astype(str))
                    match = re.search(r'\d{8}-\d{8}', row_str)
                    if match:
                        found_date = match.group()
                        break
                date_dict[f] = found_date or "Tarih Bulunamadi"
        except Exception:
            pass

    if not files:
        print("HATA: Gecerli veri dosyasi (# hole-pool iceren xlsx) bulunamadi.")
        sys.exit(1)

    print(f"   OK {len(files)} adet veri dosyasi tespit edildi.")
    unique_dates = set(date_dict.values())
    if len(unique_dates) > 1:
        print(f"   UYARI: Tarih farkliligi var: {unique_dates}")
    elif unique_dates:
        print(f"   OK Tarih: {list(unique_dates)[0]}\n")

    # 2. Verileri oku ve funnel tipine gore ayristir
    print("2. Veriler okunuyor...")
    data_android = {} # { funnel_type: df }
    data_ios     = {}

    for f in files:
        funnel, df_and, df_ios = read_platform_data(f)
        print(f"   {f} -> {funnel}")
        if funnel not in data_android: data_android[funnel] = []
        if funnel not in data_ios:     data_ios[funnel]     = []
        data_android[funnel].append(df_and)
        data_ios[funnel].append(df_ios)

    # Aynı funnel tipindeki dosyalari birlestir (Eger varsa)
    for ftype in list(data_android.keys()):
        data_android[ftype] = safe_concat(data_android[ftype])
    for ftype in list(data_ios.keys()):
        data_ios[ftype]     = safe_concat(data_ios[ftype])

    # 3. In-App Purchase USD verileri (CSV)
    print("3. CSV Gelir verileri okunuyor...")
    csv_and_usd = pd.DataFrame()
    csv_ios_usd = pd.DataFrame()
    csv_files = glob.glob("veri-*.csv")
    csv_list = []
    for cf in csv_files:
        try:
            cf_df = pd.read_csv(cf)
            if all(k in cf_df.columns for k in ['Level', 'USD Price', 'Store Name']):
                csv_list.append(cf_df)
        except Exception: pass

    if csv_list:
        all_csv = pd.concat(csv_list, ignore_index=True)
        all_csv['USD Price'] = pd.to_numeric(all_csv['USD Price'], errors='coerce').fillna(0)
        all_csv['Level']     = pd.to_numeric(all_csv['Level'], errors='coerce').dropna().astype(int)
        
        gu = all_csv[all_csv['Store Name'].astype(str).str.contains('Google', case=False, na=False)]
        csv_and_usd = gu.groupby('Level')['USD Price'].sum().reset_index().rename(columns={'USD Price': 'Total $'})
        
        ap = all_csv[all_csv['Store Name'].astype(str).str.contains('AppStore', case=False, na=False)]
        csv_ios_usd = ap.groupby('Level')['USD Price'].sum().reset_index().rename(columns={'USD Price': 'Total $'})
        print(f"   OK Gelir verileri okundu.")

    # 4. Yazma Islemi
    print("4. Sablona yaziliyor...")
    template_files = glob.glob("Copy of Hole_Pool*.xlsx")
    if not template_files:
        print("HATA: Sablon bulunamadi.")
        sys.exit(1)

    template_path = template_files[0]
    output_path   = "Hole_Pool_Otomatik_Analiz_v1.xlsx"

    try:
        wb = openpyxl.load_workbook(template_path)
    except Exception as e:
        print(f"HATA: Sablon yuklenemedi: {e}")
        sys.exit(1)

    # Genel sekmeler
    write_tab_data(wb, 'And', data_android, GENEL_COL_MAP, max_level, csv_and_usd)
    write_tab_data(wb, 'IOS', data_ios,     GENEL_COL_MAP, max_level, csv_ios_usd)

    # Booster sekmeleri
    write_tab_data(wb, 'Boosters_And', data_android, BOOSTER_COL_MAP, max_level)
    write_tab_data(wb, 'Boosters_IOS', data_ios,     BOOSTER_COL_MAP, max_level)

    print("\n5. Sonuc dosyasi kaydediliyor...")
    wb.save(output_path)
    print("=======================================================")
    print(f" BASARILI: {output_path}")
    print("=======================================================\n")

if __name__ == '__main__':
    main()
