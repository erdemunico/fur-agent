# -*- coding: utf-8 -*-
import pandas as pd
import openpyxl
from openpyxl.descriptors.base import Set
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
import glob
import sys
import warnings
import re
from datetime import datetime
import errno
import os
import json
import time
from copy import deepcopy

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Calisma kitabi sablonu (SCRIPT_DIR'de tam bu dosya adi; Git'te izlenir)
WORKBOOK_TEMPLATE_XLSX = (
    "Hole_Pool_Otomatik_Analiz_v1_US_20260325-20260331_20260401_212238.xlsx"
)

# ARPU / Av.Rw grafikleri bu referans dosyadan kopyalanir (U,V,W -> And!O / IOS!O, And!T / IOS!T)
ARPU_CHART_REFERENCE_XLSX = "Hole_Pool_Otomatik_Analiz_v1_US_20260319-20260325.xlsx"

# #region agent log
_DEBUG_LOG_PATH = os.path.join(SCRIPT_DIR, "debug-fb19ec.log")
_DEBUG_SESSION = "fb19ec"


def _agent_debug_log(hypothesis_id, location, message, data=None, run_id="run1"):
    try:
        rec = {
            "sessionId": _DEBUG_SESSION,
            "runId": run_id,
            "hypothesisId": hypothesis_id,
            "location": location,
            "message": message,
            "data": data or {},
            "timestamp": int(time.time() * 1000),
        }
        with open(_DEBUG_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
    except Exception:
        pass


# #endregion


def _is_output_workbook(basename: str) -> bool:
    b = str(basename).lower()
    return "_otomatik_analiz" in b


def _is_template_workbook(basename: str) -> bool:
    b = str(basename)
    if b.startswith("Copy of Hole_Pool"):
        return True
    return b.casefold() == WORKBOOK_TEMPLATE_XLSX.casefold()


def _normalize_game_name_for_filename(game_name: str) -> str:
    """
    Oyun adini dosya-adina uygun hale getirir.
    Ornek: 'Farm Block Escape' -> 'Farm_Block_Escape'
    """
    s = str(game_name or "").strip()
    if not s:
        return "Hole_Pool"
    s = re.sub(r"[^A-Za-z0-9]+", "_", s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s or "Hole_Pool"


def _extract_game_name_from_metadata(df_check):
    """
    Metaveriden oyun adini cikarir (oncelik: hashtag satiri).
    Ornekler:
    - '# hole-pool' -> 'hole-pool'
    - '# FARM BLOCK ESCAPE' -> 'FARM BLOCK ESCAPE'
    """
    if df_check is None or df_check.empty:
        return None
    head = df_check.head(min(15, len(df_check)))
    blob = " ".join(head.fillna("").astype(str).values.flatten())

    m = re.search(r"#\s*([A-Za-z0-9][A-Za-z0-9 _-]{1,80})", blob)
    if m:
        raw = m.group(1).strip(" -_")
        if raw:
            return raw
    return None


def _xlsx_rows_look_like_data(df_check) -> bool:
    """# hole-pool etiketi veya Free form tarzi (OS + Level + funnel) yapi."""
    if df_check is None or df_check.empty:
        return False
    # 1) Once metaveri kontrolu (oyun etiketi)
    if _extract_game_name_from_metadata(df_check):
        return True
    content_str = " ".join(df_check.fillna("").astype(str).values.flatten()).lower()
    if "# hole-pool" in content_str:
        return True
    # 2) Metaveri yoksa free-form kolon yapisi
    head = df_check.head(min(15, len(df_check)))
    blob = " ".join(head.fillna("").astype(str).values.flatten()).lower()
    if "operating system" not in blob or "level" not in blob:
        return False
    return any(k in blob for k in ("start", "complete", "fail", "tryagain", "ads"))


def collect_data_xlsx_files(script_dir):
    """
    Her zaman process_data.py'nin bulundugu klasordeki .xlsx dosyalarini tarar
    (calisma dizininden bagimsiz). Cikti ve sablon dosyalarini atlar.
    """
    all_paths = glob.glob(os.path.join(script_dir, "*.xlsx"))
    all_paths.sort(key=lambda p: (os.path.getmtime(p), p), reverse=True)

    files = []
    date_dict = {}
    game_dict = {}
    skipped = []

    for path in all_paths:
        base = os.path.basename(path)
        if _is_template_workbook(base):
            skipped.append((base, "sablon"))
            continue
        if _is_output_workbook(base):
            skipped.append((base, "cikti"))
            continue
        try:
            df_check = pd.read_excel(path, nrows=50, header=None)
        except Exception as e:
            skipped.append((base, f"okunamadi: {e!s}"))
            continue
        if not _xlsx_rows_look_like_data(df_check):
            skipped.append((base, "veri yapisi (hole-pool / OS+Level) yok"))
            continue
        files.append(path)
        game_name = _extract_game_name_from_metadata(df_check)
        game_dict[path] = game_name
        found_date = None
        for idx, row in df_check.iterrows():
            row_str = " ".join(row.dropna().astype(str))
            match = re.search(r"\d{8}-\d{8}", row_str)
            if match:
                found_date = match.group()
                break
        date_dict[path] = found_date or "Tarih Bulunamadi"

    return files, date_dict, game_dict, skipped


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
    ('Total users',     'ADs'):      20, # ham ADs; T (20) Python ile rewarded/complete orani yazilir
    # U (21) avg. owned coin: Complete coin_owned / Complete Active users — write_and_ios_t_u_columns
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


def get_arpu_chart_max_level_from_user(main_max_level):
    """
    ARPU / Av.Rw sekmelerinde grafik ve U,V,W tablosu icin ust level.
    Ana analizden kucuk veya esit olmali (And/IOS'ta veri olmayan satira baglanmaz).
    """
    prompt = (
        "ARPU / Av.Rw grafikleri kacinci level'e kadar olsun? "
        f"(Enter = ana analiz ile ayni [{main_max_level}]; ornek sadece 250: 250) : "
    )
    while True:
        try:
            raw = input(prompt).strip()
            if not raw:
                return main_max_level
            v = int(raw)
            if v < 1:
                print("   En az 1 girin.")
                continue
            if v > main_max_level:
                print(
                    f"   UYARI: Ana analiz {main_max_level} level; grafik limiti "
                    f"{main_max_level} ile sinirlandi."
                )
                return main_max_level
            return v
        except ValueError:
            print("   Gecerli bir tam sayi girin veya Enter'a basin.")
        except EOFError:
            return main_max_level


def _base_col_name(col_name):
    if "_" in col_name and str(col_name).split("_")[-1].isdigit():
        return col_name.split("_")[0]
    return col_name


def _norm_metric_token(s):
    return str(s).strip().lower().replace(" ", "_").replace("-", "_")


# Android export siklikla "Owned Coin" / owned_coin; iOS coin_owned — U sutunu icin ayni metrik.
_COIN_CANON = "coin_owned"
_COIN_ALIASES = frozenset(
    {"coin_owned", "owned_coin", "coinowned", "ownedcoin", "coins_owned", "owned_coins"}
)


def _canonical_metric_name(name):
    """Sutun adini karsilastirma icin tek forma getir (And/IOS baslik farki)."""
    k = _norm_metric_token(_base_col_name(name))
    if k in _COIN_ALIASES:
        return _COIN_CANON
    return k


def _canonical_target_base(target_base):
    k = _norm_metric_token(target_base)
    if k in _COIN_ALIASES:
        return _COIN_CANON
    return k


def get_metric_at_level(processed_by_ftype, level, ftype, target_base):
    """Funnel df icinde hedef metrik; sutun adi Android/iOS'ta farkli yazilabilir."""
    df = processed_by_ftype.get(ftype)
    if df is None or df.empty or level not in df.index:
        return 0.0
    row = df.loc[level]
    want = _canonical_target_base(target_base)
    for cname in df.columns:
        if _canonical_metric_name(cname) == want:
            v = pd.to_numeric(row[cname], errors="coerce")
            return float(v) if pd.notna(v) else 0.0
    return 0.0


def build_processed_by_funnel(data_by_funnel, max_level):
    out = {}
    for ftype in ["Start", "Complete", "Fail", "Tryagain", "ADs"]:
        df = data_by_funnel.get(ftype, pd.DataFrame())
        out[ftype] = format_and_fill_levels(df, max_level).set_index("Level")
    return out


def _worksheet_by_name(wb, logical_name):
    """Excel sekmesi adi And/AND/Android veya IOS farkli yazilabilsin."""
    if logical_name in wb.sheetnames:
        return wb[logical_name]
    low = logical_name.strip().lower()
    for n in wb.sheetnames:
        if n.strip().lower() == low:
            return wb[n]
    if low == "and":
        for n in wb.sheetnames:
            if n.strip().lower() == "android":
                return wb[n]
    return None


def write_and_ios_t_u_columns(wb, sheet_name, data_by_funnel, max_level):
    """
    And / IOS:
    - T (20): ADs Total users / Complete Active users (ortalama rewarded).
    - U (21): Complete coin_owned / Complete Active users (avg. owned coin).
    - U1: Yalnizca sablonda 'U1' basligi varsa; U sutununun kademeli ortalamasi (X varsayilan yok).
    V (22) ve V1 (23): write_tab_data ile sablon formulleri.
    """
    if sheet_name not in ("And", "IOS"):
        return
    ws = _worksheet_by_name(wb, sheet_name)
    if ws is None:
        print(f"   ! Uyari: '{sheet_name}' sekmesi bulunamadi (T/U atlaniyor).")
        return
    u1_col = resolve_u1_column_for_sheet(ws)
    level_rows = find_level_rows(ws)
    proc = build_processed_by_funnel(data_by_funnel, max_level)

    for level in range(1, max_level + 1):
        complete = get_metric_at_level(proc, level, "Complete", "Active users")
        rewarded = get_metric_at_level(proc, level, "ADs", "Total users")
        coin = get_metric_at_level(proc, level, "Complete", "coin_owned")
        if complete and complete != 0:
            t_val = rewarded / complete
            u_val = coin / complete
        else:
            t_val = 0.0
            u_val = 0.0
        t_row = level_rows.get(level, level + 2)
        ws.cell(row=t_row, column=20).value = t_val
        ws.cell(row=t_row, column=21).value = u_val

    if u1_col is not None and u1_col not in (20, 21):
        first_u_row = level_rows.get(1, 3)
        for level in range(1, max_level + 1):
            t_row = level_rows.get(level, level + 2)
            ws.cell(row=t_row, column=u1_col).value = (
                f"=AVERAGE($U${first_u_row}:U{t_row})"
            )

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


def _is_u1_header_value(v):
    """Hucre metni U1 basligi mi (tam veya 'U1 ortalama' gibi)."""
    if v is None:
        return False
    s = str(v).strip()
    if not s:
        return False
    return bool(re.match(r"^U1\b", s, re.IGNORECASE))


def _find_u1_header_column(ws):
    """Sablonda U1 basligi (satir 1-3). T (20) / U (21) atlanir."""
    max_c = ws.max_column or 0
    for r in (1, 2, 3):
        for c in range(1, max_c + 1):
            if c in (20, 21):
                continue
            v = ws.cell(row=r, column=c).value
            if _is_u1_header_value(v):
                return c
    return None


def resolve_u1_column_for_sheet(ws):
    """
    Her sekme kendi U1 sutununu kullanir.
    Baslik yoksa None — X (24) veya baska sutun otomatik doldurulmaz (grafik / sablon korunur).
    """
    return _find_u1_header_column(ws)


def _norm_header_cell_text(v):
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v).strip().lower()).replace("_", " ")


def _find_column_by_header_keywords(ws, keywords, rows=(1, 2, 3), skip_cols=None):
    """Satir 1-3'te hucre metni keywords'den birini iceren ilk sutun."""
    skip_cols = skip_cols or frozenset()
    max_c = ws.max_column or 0
    for r in rows:
        for c in range(1, max_c + 1):
            if c in skip_cols:
                continue
            nh = _norm_header_cell_text(ws.cell(row=r, column=c).value)
            if not nh:
                continue
            for kw in keywords:
                if kw in nh:
                    return c
    return None


def write_row1_owned_avg_coin_averages(wb, sheet_name, max_level):
    """
    And / IOS: 1. satira Owned Coin ve Avg. Coin sutunlarinin
    level veri satirlari (1..max_level) uzerinden AVERAGE formulu.
    """
    if sheet_name not in ("And", "IOS"):
        return
    ws = _worksheet_by_name(wb, sheet_name)
    if ws is None:
        return
    level_rows = find_level_rows(ws)
    if not level_rows or max_level < 1:
        return
    first_row = level_rows.get(1, 3)
    last_row = level_rows.get(max_level, max_level + 2)

    col_owned = _find_column_by_header_keywords(ws, ("owned coin", "coin owned"))
    skip = frozenset([col_owned]) if col_owned is not None else frozenset()
    col_avg = _find_column_by_header_keywords(
        ws,
        ("avg. coin", "avg coin", "average coin"),
        skip_cols=skip,
    )
    if col_owned is None:
        col_owned = 21
    if col_avg is None:
        col_avg = 22
    if col_avg == col_owned:
        col_avg = 22

    lo = get_column_letter(col_owned)
    lv = get_column_letter(col_avg)
    ws.cell(row=1, column=col_owned).value = (
        f"=AVERAGE({lo}{first_row}:{lo}{last_row})"
    )
    ws.cell(row=1, column=col_avg).value = (
        f"=AVERAGE({lv}{first_row}:{lv}{last_row})"
    )


def find_arpu_reference_workbook(script_dir, output_path=None):
    """
    Oncelik: ARPU_CHART_REFERENCE_XLSX. Yoksa en yeni Hole_Pool_Otomatik_Analiz*.xlsx
    (cikti dosyasi ve sablon haric).
    """
    out_abs = os.path.abspath(output_path) if output_path else None
    pref = os.path.join(script_dir, ARPU_CHART_REFERENCE_XLSX)
    if os.path.isfile(pref) and (out_abs is None or os.path.abspath(pref) != out_abs):
        return pref
    paths = glob.glob(os.path.join(script_dir, "Hole_Pool_Otomatik_Analiz*.xlsx"))
    paths = [
        p
        for p in paths
        if not _is_template_workbook(os.path.basename(p))
        and (out_abs is None or os.path.abspath(p) != out_abs)
    ]
    paths.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return paths[0] if paths else None


def _copy_sheet_cells_and_merges(src_ws, dst_ws):
    for row in src_ws.iter_rows():
        for c in row:
            dst_ws.cell(row=c.row, column=c.column, value=c.value)
    for rng in list(src_ws.merged_cells.ranges):
        dst_ws.merge_cells(str(rng))


def copy_arpu_avgrw_sheets_from_reference(dst_wb, ref_path):
    """Referanstaki ARPU ve Av.Rw sekmelerini (cizgi grafikleriyle) hedef calisma kitabina alir."""
    ref_wb = openpyxl.load_workbook(ref_path, rich_text=True, data_only=False)
    try:
        for name in ("ARPU", "Av.Rw"):
            if name not in ref_wb.sheetnames:
                raise ValueError(f"Referansta '{name}' sekmesi yok.")
            if name in dst_wb.sheetnames:
                dst_wb.remove(dst_wb[name])
        for insert_idx, name in enumerate(("ARPU", "Av.Rw")):
            src = ref_wb[name]
            dst = dst_wb.create_sheet(name, insert_idx)
            _copy_sheet_cells_and_merges(src, dst)
            for col_letter, dim in src.column_dimensions.items():
                if dim.width:
                    dst.column_dimensions[col_letter].width = dim.width
            dst._charts = [deepcopy(ch) for ch in (src._charts or [])]
    finally:
        ref_wb.close()


def _excel_quoted_sheet_title(wb, logical_name):
    ws = _worksheet_by_name(wb, logical_name)
    t = ws.title if ws else logical_name
    return "'" + str(t).replace("'", "''") + "'"


def _replace_series_range_end_row(ref_formula, last_row):
    if not ref_formula or not isinstance(ref_formula, str):
        return ref_formula

    def _repl(m):
        return m.group(1) + str(last_row)

    return re.sub(r"(:\$[A-Z]+\$)\d+$", _repl, ref_formula)


def _stretch_line_chart_for_max_level(chart, last_row):
    for ser in chart.series:
        if getattr(ser, "val", None) and ser.val.numRef and ser.val.numRef.f:
            ser.val.numRef.f = _replace_series_range_end_row(ser.val.numRef.f, last_row)
        cat = getattr(ser, "cat", None)
        if not cat:
            continue
        if cat.strRef and cat.strRef.f:
            cat.strRef.f = _replace_series_range_end_row(cat.strRef.f, last_row)
        elif getattr(cat, "numRef", None) and cat.numRef and cat.numRef.f:
            cat.numRef.f = _replace_series_range_end_row(cat.numRef.f, last_row)


def apply_arpu_avgrw_level_formulas_and_charts(wb, chart_max_level):
    """
    ARPU: U=level, V=And!O (ARPU), W=IOS!O. Av.Rw: U=level, V=And!T, W=IOS!T.
    Satir 3 = level 1; son veri satiri = chart_max_level + 2.
    Grafik seri araliklari buna gore; ustundeki U,V,W temizlenir.
    """
    if chart_max_level < 1:
        return
    last_row = chart_max_level + 2
    and_ref = _excel_quoted_sheet_title(wb, "And")
    ios_ref = _excel_quoted_sheet_title(wb, "IOS")

    if "ARPU" in wb.sheetnames:
        ws = wb["ARPU"]
        for r in range(3, last_row + 1):
            lvl = r - 2
            ws.cell(row=r, column=21, value=float(lvl))
            ws.cell(row=r, column=22, value=f"={and_ref}!O{r}")
            ws.cell(row=r, column=23, value=f"={ios_ref}!O{r}")
        for ch in ws._charts or []:
            _stretch_line_chart_for_max_level(ch, last_row)

    if "Av.Rw" in wb.sheetnames:
        ws = wb["Av.Rw"]
        for r in range(3, last_row + 1):
            lvl = r - 2
            ws.cell(row=r, column=21, value=float(lvl))
            ws.cell(row=r, column=22, value=f"={and_ref}!T{r}")
            ws.cell(row=r, column=23, value=f"={ios_ref}!T{r}")
        for ch in ws._charts or []:
            _stretch_line_chart_for_max_level(ch, last_row)

    # Referansta kalan asiri satirlari grafik disinda bosalt (eski formuller kalmasin)
    for name in ("ARPU", "Av.Rw"):
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        for r in range(last_row + 1, 1001):
            for c in (21, 22, 23):
                ws.cell(row=r, column=c).value = None


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


# Kod analizi: Drop 1 (H) ve % / K sütunu (K) — formüller Excel'de hesaplanır; renk CF ile.
# Drop 1: >%10 kirmizi, %8.5–10 (dahil) sari.
# K (%): Level araliklarina gore parca parca lineer esikler
# - 20  : yesil>50, sari 45-50
# - 30  : yesil>30, sari 25-30
# - 50  : yesil>15, sari 12-15
# - 100 : yesil>7,  sari 5-7
# Not: 100 esikleri 50'den sonra baslar (51+). 100+ level'da K sutununu bos birakiyoruz.
_CF_FILL_GREEN = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
_CF_FILL_YELLOW = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
_CF_FILL_RED = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

# Excel IF zinciri (K esik degerleri):
# g: yesil esigi (ustu yesil)
# y: sari bandin alt esigi (y ile g arasi sari)
_K_THRESH_G = (
    'IF($A3<=20,50,IF($A3<=30,50-2*($A3-20),IF($A3<=50,30-0.75*($A3-30),'
    'IF($A3<=100,15-0.16*($A3-50),7))))'
)
_K_THRESH_Y = (
    'IF($A3<=20,45,IF($A3<=30,45-2*($A3-20),IF($A3<=50,25-0.65*($A3-30),'
    'IF($A3<=100,12-0.14*($A3-50),5))))'
)


def apply_drop1_and_k_conditional_formatting(ws, last_row):
    """And / IOS: H (Drop 1) ve K (%) için koşullu dolgu. Formüller hesaplandıktan sonra Excel uygular."""
    if last_row < 3:
        return
    h_range = f'H3:H{last_row}'
    k_range = f'K3:K{last_row}'
    # Drop 1: sadece sayı varsa renklendir; boş/hata -> beyaz (varsayılan)
    ws.conditional_formatting.add(
        h_range,
        FormulaRule(formula=['AND(ISNUMBER(H3),H3>10)'], fill=_CF_FILL_RED, stopIfTrue=True),
    )
    ws.conditional_formatting.add(
        h_range,
        FormulaRule(
            formula=['AND(ISNUMBER(H3),H3>=8.5,H3<=10)'],
            fill=_CF_FILL_YELLOW,
            stopIfTrue=True,
        ),
    )
    # K (%): sadece hem Level (A3) hem K sayisal ise renk; bos/metin -> renk yok
    ws.conditional_formatting.add(
        k_range,
        FormulaRule(
            formula=[f'AND(ISNUMBER($A3),ISNUMBER(K3),K3>({_K_THRESH_G}))'],
            fill=_CF_FILL_GREEN,
            stopIfTrue=True,
        ),
    )
    ws.conditional_formatting.add(
        k_range,
        FormulaRule(
            formula=[
                f'AND(ISNUMBER($A3),ISNUMBER(K3),K3>=({_K_THRESH_Y}),K3<=({_K_THRESH_G}))'
            ],
            fill=_CF_FILL_YELLOW,
            stopIfTrue=True,
        ),
    )
    ws.conditional_formatting.add(
        k_range,
        FormulaRule(
            formula=[
                f'AND(ISNUMBER($A3),ISNUMBER(K3),K3<({_K_THRESH_Y}))'
            ],
            fill=_CF_FILL_RED,
            stopIfTrue=True,
        ),
    )


def write_tab_data(wb, sheet_name, data_by_funnel, col_map, max_level, revenue_df=None):
    """
    Belirli bir sekme icin verileri funnel haritasina gore yazdirir.
    Level 1'den max_level'a kadar tüm satırları doldurur.
    """
    if sheet_name in ("And", "IOS"):
        ws = _worksheet_by_name(wb, sheet_name)
        if ws is None:
            print(f"   ! Uyari: '{sheet_name}' sekmesi yok, atlaniyor.")
            return
    else:
        if sheet_name not in wb.sheetnames:
            print(f"   ! Uyari: '{sheet_name}' sekmesi yok, atlaniyor.")
            return
        ws = wb[sheet_name]

    u1_col = resolve_u1_column_for_sheet(ws) if sheet_name in ("And", "IOS") else None

    # 1. TEMIZLIK: Sadece veri yazilacak sutunlari temizle (Row 3-1000)
    cols_to_clear = set(col_map.values())
    if sheet_name in ['And', 'IOS']:
        cols_to_clear.add(14) # Total $
        cols_to_clear.add(21) # U: avg. owned coin (Python)
        if u1_col is not None and u1_col not in (20, 21):
            cols_to_clear.add(u1_col)
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
                        if sheet_name in ['And', 'IOS'] and col_idx == 20:
                            continue  # T: Python (write_and_ios_t_u_columns)
                        if sheet_name in ['And', 'IOS'] and col_idx == 21:
                            continue  # U: avg. owned coin — Python
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
        if sheet_name in ['And', 'IOS'] and c == 20:
            continue  # T: Python
        if sheet_name in ['And', 'IOS'] and c == 21:
            continue  # U: Python
        if sheet_name in ['And', 'IOS'] and c == 23:
            continue  # V1: sablon
        if sheet_name in ['And', 'IOS'] and u1_col is not None and c == u1_col:
            continue  # U1: write_and_ios_t_u_columns

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

    # 5b. K (%) sütunu:
    # - Level 1..100: şablondaki formül yazılır
    # - Level >100 : boş bırakılır (gösterilmez)
    # Formül: L1 Complete (satır 3, C3) ile bir alt satırdaki C sütunu arasındaki oran; yüzde cinsinden.
    if sheet_name in ('And', 'IOS'):
        for level in range(1, max_level + 1):
            t_row = level_rows[level]
            if level <= 100:
                ws.cell(row=t_row, column=11).value = (
                    f'=100-(((C3-C{t_row + 1})/C3)*100)'
                )
            else:
                ws.cell(row=t_row, column=11).value = None

    if sheet_name in ('And', 'IOS'):
        # Sadece bizim doldurduğumuz max_level araligina kadar renklendirme yap
        # (template'te daha uzun seviyeler tanimli olsa bile).
        last_row_fmt = level_rows.get(max_level, max_level + 2)
        apply_drop1_and_k_conditional_formatting(ws, last_row_fmt)

    print(f"   OK '{sheet_name}' tablosu yazildi (Level 1-{max_level}).")


def main():
    # #region agent log
    _agent_debug_log(
        "H1",
        "process_data.py:main:entry",
        "main started (syntax OK, file parsed)",
        {},
        "run1",
    )
    # #endregion
    print("=======================================================")
    print(" Hole Pool Veri Isleme Scripti Baslatiliyor...")
    print("=======================================================\n")

    max_level = get_max_level_from_user()
    print(f"\n[OK] Analiz {max_level}. levele kadar yapilacak.\n")

    # 1. Veri dosyalarini SCRIPT_DIR icinde tara (cwd'den bagimsiz)
    print("1. Veri dosyalari (xlsx) taraniyor...")
    print(f"   Klasor: {os.path.normpath(SCRIPT_DIR)}")
    files, date_dict, game_dict, skipped = collect_data_xlsx_files(SCRIPT_DIR)

    if not files:
        print("HATA: Gecerli veri dosyasi bulunamadi (hole-pool etiketi veya OS+Level+funnel yapisinda xlsx).")
        if skipped:
            print("   Atlanan dosyalar (ilk 20):")
            for name, why in skipped[:20]:
                print(f"     - {name}: {why}")
            if len(skipped) > 20:
                print(f"     ... ve {len(skipped) - 20} dosya daha")
        sys.exit(1)

    if skipped:
        print(f"   Bilgi: {len(skipped)} dosya atlandi (sablon/cikti veya veri degil).")
    print(f"   OK {len(files)} adet veri dosyasi tespit edildi.")
    game_names = [g for g in game_dict.values() if g]
    if game_names:
        # Birden fazla oyun etiketi varsa en sik geceni baz al.
        game_counter = {}
        for g in game_names:
            game_counter[g] = game_counter.get(g, 0) + 1
        detected_game = sorted(game_counter.items(), key=lambda x: (-x[1], x[0]))[0][0]
        print(f"   OK Oyun metaverisi: {detected_game}")
    else:
        detected_game = "Hole_Pool"
        print("   UYARI: Oyun metaverisi bulunamadi; varsayilan ad kullanilacak (Hole_Pool).")
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
        print(f"   {os.path.basename(f)} -> {funnel}")
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
    csv_files = glob.glob(os.path.join(SCRIPT_DIR, "veri-*.csv"))
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
    template_path = os.path.join(SCRIPT_DIR, WORKBOOK_TEMPLATE_XLSX)
    if not os.path.isfile(template_path):
        print(
            f"HATA: Sablon bulunamadi: {WORKBOOK_TEMPLATE_XLSX}\n"
            f"      Beklenen konum: {os.path.normpath(template_path)}"
        )
        sys.exit(1)

    # --- Output dosya adi: ulke kisaltmasi + analiz tarih araligi ---
    # unique_dates daha once, #hole-pool iceren dosyalardan cikartiliyordu.
    # unique_dates birden fazla olabilir; hepsini kapsayacak sekilde min-max aliyoruz.
    if unique_dates:
        ranges = []
        for dr in unique_dates:
            parts = str(dr).split("-")
            if len(parts) == 2 and all(p.isdigit() for p in parts):
                ranges.append((parts[0], parts[1]))
        if ranges:
            start = min(r[0] for r in ranges)
            end = max(r[1] for r in ranges)
            date_range = f"{start}-{end}"
        else:
            date_range = str(list(unique_dates)[0])
    else:
        date_range = "TarihBulunamadi"
    date_range_safe = str(date_range).replace(" ", "")

    def normalize_country_code(x: str) -> str:
        x = str(x).upper().strip()
        # Kullanici ornegi: USA -> US
        m = {
            "USA": "US",
            "UNITEDSTATES": "US",
        }
        return m.get(x, x)

    country_code = "XX"
    tmpl_bn = os.path.basename(template_path)
    m = re.search(r"Hole_Pool_Otomatik_Analiz_v1_([A-Za-z]+)_", tmpl_bn, re.I)
    if m:
        country_code = normalize_country_code(m.group(1))
    else:
        m2 = re.search(r"Hole_Pool_([A-Za-z]+)_", tmpl_bn)
        if m2:
            country_code = normalize_country_code(m2.group(1))
        else:
            m3 = re.search(r"Hole_Pool_([A-Za-z]+)", tmpl_bn)
            if m3:
                country_code = normalize_country_code(m3.group(1))

    game_slug = _normalize_game_name_for_filename(detected_game)
    output_path = os.path.join(
        SCRIPT_DIR,
        f"{game_slug}_Otomatik_Analiz_v1_{country_code}_{date_range_safe}.xlsx",
    )

    try:
        # rich_text=True: grafik / sekme ogesi ile iliskili zengin metin ve round-trip icin daha iyi korunur
        wb = openpyxl.load_workbook(template_path, rich_text=True)
    except Exception as e:
        print(f"HATA: Sablon yuklenemedi: {e}")
        sys.exit(1)

    ref_arpu_path = find_arpu_reference_workbook(SCRIPT_DIR, output_path)
    if ref_arpu_path:
        try:
            copy_arpu_avgrw_sheets_from_reference(wb, ref_arpu_path)
            print(
                f"   OK ARPU/Av.Rw (grafikler) referanstan kopyalandi: "
                f"{os.path.basename(ref_arpu_path)}"
            )
        except Exception as e:
            print(f"   UYARI: ARPU/Av.Rw referansi kopyalanamadi: {e}")

    arpu_chart_max_level = max_level
    if "ARPU" in wb.sheetnames and "Av.Rw" in wb.sheetnames:
        print()
        arpu_chart_max_level = get_arpu_chart_max_level_from_user(max_level)
        print(f"[OK] ARPU / Av.Rw grafikleri {arpu_chart_max_level}. levele kadar.\n")

    # Excel'in kayit dosyasini acarken formulleri hesaplayip CF degerlendirmesini yapmasi icin.
    try:
        if getattr(wb, "calculation", None) is not None:
            wb.calculation.fullCalcOnLoad = True
    except Exception:
        pass

    # Booster sekmeleri: And/IOS'tan once yazilir (T/V1 icin sayfa sirasi)
    write_tab_data(wb, 'Boosters_And', data_android, BOOSTER_COL_MAP, max_level)
    write_tab_data(wb, 'Boosters_IOS', data_ios, BOOSTER_COL_MAP, max_level)

    # Genel sekmeler
    write_tab_data(wb, 'And', data_android, GENEL_COL_MAP, max_level, csv_and_usd)
    write_tab_data(wb, 'IOS', data_ios,     GENEL_COL_MAP, max_level, csv_ios_usd)

    write_and_ios_t_u_columns(wb, 'And', data_android, max_level)
    write_and_ios_t_u_columns(wb, 'IOS', data_ios, max_level)

    write_row1_owned_avg_coin_averages(wb, "And", max_level)
    write_row1_owned_avg_coin_averages(wb, "IOS", max_level)

    if "ARPU" in wb.sheetnames and "Av.Rw" in wb.sheetnames:
        apply_arpu_avgrw_level_formulas_and_charts(wb, arpu_chart_max_level)

    # #region agent log
    _agent_debug_log(
        "H2",
        "process_data.py:main:before_save",
        "write order done; about to save",
        {
            "max_level": max_level,
            "arpu_chart_max_level": arpu_chart_max_level,
            "output_path": output_path,
        },
        "run1",
    )
    # #endregion

    print("\n5. Sonuc dosyasi kaydediliyor...")
    try:
        print(f"   -> Kaydediliyor: {output_path}")
        wb.save(output_path)
    except Exception as e:
        # Windows'ta çıktı dosyası Excel'de açıkken yazma kilitlenir (errno 13).
        if not (isinstance(e, PermissionError) or getattr(e, "errno", None) == errno.EACCES):
            raise
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        alt = output_path.replace(".xlsx", f"_{ts}.xlsx")
        print(
            f"UYARI: '{output_path}' yazilamadi (dosya Excel'de acik olabilir — kapatip tekrar deneyin).\n"
            f"        Alternatif dosyaya yaziliyor: {alt}"
        )
        wb.save(alt)
        output_path = alt
    print("=======================================================")
    print(f" BASARILI: {output_path}")
    print("=======================================================\n")
    # #region agent log
    _agent_debug_log(
        "H3",
        "process_data.py:main:success",
        "save completed",
        {"final_output_path": output_path},
        "run1",
    )
    # #endregion

if __name__ == '__main__':
    main()
