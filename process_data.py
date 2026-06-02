# -*- coding: utf-8 -*-
import pandas as pd
import openpyxl
from openpyxl.descriptors.base import Set
from openpyxl.styles import Alignment, Font, PatternFill
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

# Ozet sekmesi: erken level / segmentasyon kalite analizi
SUMMARY_SHEET_NAME = "Ozet"
QUALITY_MIN_USERS = 10
QUALITY_COVERAGE_RATIO = 0.05
QUALITY_CONSECUTIVE_LEVELS = 3
QUALITY_REF_WINDOW = 40

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
    ('Total users',     'Start'):    2,
    ('Active users',    'Complete'): 3,
    ('Total users',     'Complete'): 3,
    ('Active users',    'Fail'):     4,
    ('Total users',     'Fail'):     4,
    ('Active users',    'Tryagain'): 5,
    ('Total users',     'Tryagain'): 5,
    ('avarage_attempt', 'Complete'): 12,
    ('avg_thinktime',   'Complete'): 19,
    ('Total users',     'ADs'):      20,  # ham ADs; T (20) Python ile rewarded/complete orani
    ('Active users',    'ADs'):      20,
    # U (21) avg. owned coin: Complete coin_owned / kullanici sayisi — write_and_ios_t_u_columns
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
    ('Total users',     'ADs'):      21,  # Total Rewarded
    ('Active users',    'ADs'):      21,
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
_USER_COUNT_CANON = "user_count"
_USER_COUNT_ALIASES = frozenset(
    {"active_users", "total_users", "activeusers", "totalusers"}
)

# Funnel tipi tespiti (ust satir / Segment); uzun eslesmeler once
_FUNNEL_DETECT_RULES = (
    ("try again", "Tryagain"),
    ("tryagain", "Tryagain"),
    ("complete", "Complete"),
    ("fail", "Fail"),
    ("rewarded", "ADs"),
    ("adtype", "ADs"),
    ("ads", "ADs"),
    ("start", "Start"),
)


def _canonical_metric_name(name):
    """Sutun adini karsilastirma icin tek forma getir (And/IOS baslik farki)."""
    k = _norm_metric_token(_base_col_name(name))
    if k in _COIN_ALIASES:
        return _COIN_CANON
    if k in _USER_COUNT_ALIASES:
        return _USER_COUNT_CANON
    return k


def _canonical_target_base(target_base):
    k = _norm_metric_token(target_base)
    if k in _COIN_ALIASES:
        return _COIN_CANON
    if k in _USER_COUNT_ALIASES:
        return _USER_COUNT_CANON
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
        row_str = " ".join(df_raw.iloc[idx].dropna().astype(str)).lower()
        for needle, ftype in _FUNNEL_DETECT_RULES:
            if needle in row_str:
                return ftype
    return "Unknown"


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

    if sheet_name in ("Boosters_And", "Boosters_IOS"):
        _apply_booster_pct_used_formulas(ws, sheet_name, max_level, level_rows)

    print(f"   OK '{sheet_name}' tablosu yazildi (Level 1-{max_level}).")


def _apply_booster_pct_used_formulas(ws, sheet_name, max_level, level_rows):
    """
    I sutunu (%_used_booster): her platform kendi sekmesini kullanir.
    Ornek: =(Boosters_IOS!B14*100)/Boosters_IOS!N14
    """
    if sheet_name not in ("Boosters_And", "Boosters_IOS"):
        return
    prefix = f"{sheet_name}!"
    for level in range(1, max_level + 1):
        r = level_rows.get(level, level + 2)
        ws.cell(row=r, column=9).value = f"=({prefix}B{r}*100)/{prefix}N{r}"


def _level_user_series(df, max_level):
    """Funnel DataFrame'den level -> kullanici sayisi (Active/Total users)."""
    if df is None or df.empty or "Level" not in df.columns:
        return {}
    sub = df.copy()
    sub["Level"] = pd.to_numeric(sub["Level"], errors="coerce")
    sub = sub.dropna(subset=["Level"])
    if sub.empty:
        return {}
    sub["Level"] = sub["Level"].astype(int)
    user_cols = [
        c for c in sub.columns if c != "Level" and _canonical_metric_name(c) == _USER_COUNT_CANON
    ]
    if not user_cols:
        return {}
    for c in user_cols:
        sub[c] = pd.to_numeric(sub[c], errors="coerce").fillna(0)
    grouped = sub.groupby("Level")[user_cols].sum()
    out = {}
    for lvl, row in grouped.iterrows():
        li = int(lvl)
        if li < 1 or li > max_level:
            continue
        out[li] = float(row.sum())
    return out


def _find_reliable_start_level(level_counts, ref_window=QUALITY_REF_WINDOW):
    """
    Erken level'lerde dusuk/eksik kullanici (segmentasyon artefakti) tespiti.
    Doner: (onerilen_baslangic_level, guvensiz_level_listesi, aciklama_listesi)
    """
    notes = []
    if not level_counts:
        return 1, [], ["Kullanici sayisi serisi bos."]

    sorted_lc = sorted(level_counts.items())
    early = [(l, c) for l, c in sorted_lc if l <= ref_window]
    pool = early if early else sorted_lc[: min(30, len(sorted_lc))]
    ref = max((c for _, c in pool), default=0.0)
    if ref <= 0:
        ref = max((c for _, c in sorted_lc), default=0.0)
    if ref <= 0:
        return 1, [l for l, _ in sorted_lc], ["Tum level'larda kullanici 0."]

    threshold = max(QUALITY_MIN_USERS, ref * QUALITY_COVERAGE_RATIO)
    unreliable = [l for l, c in sorted_lc if c < threshold]

    first_ok = None
    streak = 0
    for l, c in sorted_lc:
        if c >= threshold:
            streak += 1
            if streak >= QUALITY_CONSECUTIVE_LEVELS:
                first_ok = l - QUALITY_CONSECUTIVE_LEVELS + 1
                break
        else:
            streak = 0

    if first_ok is None:
        first_ok = sorted_lc[0][0]
        notes.append("Ardisik guvenilir band bulunamadi; ilk veri level'i kullanildi.")

    l1_count = level_counts.get(1, 0.0)
    if l1_count > 0 and ref > 0 and l1_count < ref * QUALITY_COVERAGE_RATIO:
        notes.append(
            f"Level 1 kullanici ({int(l1_count)}) zirve bandina gore dusuk; "
            "user property segmentasyonu erken level'leri eksik gosterebilir."
        )
    if first_ok > 1:
        notes.append(
            f"Guvenilir analiz icin Level {first_ok}+ onerilir "
            f"(Level 1-{first_ok - 1} dusuk orneklem)."
        )

    return first_ok, unreliable, notes


def _first_positive_level(level_counts):
    for l, c in sorted(level_counts.items()):
        if c > 0:
            return l
    return None


def _missing_level_gaps(level_counts, max_level):
    if not level_counts:
        return []
    lo = min(level_counts.keys())
    hi = min(max(level_counts.keys()), max_level)
    return [l for l in range(lo, hi + 1) if l not in level_counts]


def _analyze_platform_quality(data_by_funnel, max_level, platform_label):
    """Tek platform (Android/iOS) icin detayli kalite raporu."""
    funnels = ("Start", "Complete", "Fail", "Tryagain", "ADs")
    series = {ft: _level_user_series(data_by_funnel.get(ft), max_level) for ft in funnels}
    funnel_first = {ft: _first_positive_level(series[ft]) for ft in funnels}

    funnel_starts = []
    for ft in ("Start", "Complete", "Fail", "Tryagain"):
        if series[ft]:
            rel, _, _ = _find_reliable_start_level(series[ft])
            funnel_starts.append(rel)
    recommended_funnel = max(funnel_starts) if funnel_starts else 1

    recommended_ads = 1
    if series["ADs"]:
        recommended_ads, _, ads_notes = _find_reliable_start_level(series["ADs"])
    else:
        ads_notes = []

    recommended = recommended_funnel

    all_notes = []
    for ft in funnels:
        if funnel_first[ft] and funnel_first[ft] > 1:
            all_notes.append(f"{ft}: ilk veri Level {funnel_first[ft]}")

    _, _, start_notes = _find_reliable_start_level(series["Start"])
    all_notes.extend(start_notes)
    if recommended_ads > recommended_funnel:
        all_notes.append(
            f"ADs/rewarded verisi Level {recommended_ads}+ (T sutunu bu banddan sonra anlamli)."
        )
    all_notes.extend(ads_notes)

    gaps = _missing_level_gaps(series["Start"] or series["Complete"], max_level)
    if gaps and len(gaps) <= 25:
        all_notes.append(f"Eksik level araligi (ornek): {gaps[:15]}")
    elif gaps:
        all_notes.append(f"Toplam {len(gaps)} eksik level (analytics export atlama).")

    detail_cap = min(max_level, 120)
    level_rows = []
    for lvl in range(1, detail_cap + 1):
        counts = {ft: int(series[ft].get(lvl, 0)) for ft in funnels}
        reliable = lvl >= recommended_funnel and any(
            counts[ft] for ft in ("Start", "Complete", "Fail", "Tryagain")
        )
        note_parts = []
        if lvl < recommended_funnel:
            note_parts.append("Funnel dusuk guven")
        if funnel_first.get("ADs") and lvl < funnel_first["ADs"]:
            note_parts.append("ADs henuz yok")
        if not any(counts.values()):
            note_parts.append("Veri yok")
        level_rows.append(
            {
                "level": lvl,
                "counts": counts,
                "reliable": "Evet" if reliable else "Hayir",
                "note": "; ".join(note_parts) if note_parts else "",
            }
        )

    return {
        "label": platform_label,
        "recommended_level": recommended,
        "recommended_funnel_level": recommended_funnel,
        "recommended_ads_level": recommended_ads,
        "funnel_first_level": funnel_first,
        "series": series,
        "warnings": all_notes,
        "level_rows": level_rows,
    }


def _k_green_threshold(level):
    if level <= 20:
        return 50.0
    if level <= 30:
        return 50.0 - 2.0 * (level - 20)
    if level <= 50:
        return 30.0 - 0.75 * (level - 30)
    if level <= 100:
        return 15.0 - 0.16 * (level - 50)
    return 7.0


def _k_yellow_threshold(level):
    if level <= 20:
        return 45.0
    if level <= 30:
        return 45.0 - 2.0 * (level - 20)
    if level <= 50:
        return 25.0 - 0.65 * (level - 30)
    if level <= 100:
        return 12.0 - 0.14 * (level - 50)
    return 5.0


def _excel_k_pct(complete, level):
    """And/IOS K (%) sutunu ile ayni mantik: L1 Complete vs (level+1) Complete."""
    c1 = complete.get(1, 0.0)
    c_next = complete.get(level + 1, 0.0)
    if c1 <= 0:
        return None
    return 100.0 - ((c1 - c_next) / c1 * 100.0)


def _drop1_pct(complete, level):
    """Level -> level+1 Complete dususu (%). H (Drop 1) ile uyumlu."""
    c = complete.get(level, 0.0)
    c_next = complete.get(level + 1, 0.0)
    if c <= 0:
        return None
    return (c - c_next) / c * 100.0


def _classify_k(level, k_val):
    if k_val is None:
        return ""
    g = _k_green_threshold(level)
    y = _k_yellow_threshold(level)
    if k_val > g:
        return "Yesil"
    if k_val >= y:
        return "Sari"
    return "Kirmizi"


def _classify_drop1(drop_val):
    if drop_val is None:
        return ""
    if drop_val > 10:
        return "Kirmizi"
    if drop_val >= 8.5:
        return "Sari"
    return "OK"


def _churn_priority_score(k_cls, d_cls):
    if d_cls == "Kirmizi":
        return 3
    if k_cls == "Kirmizi":
        return 2
    if d_cls == "Sari" or k_cls == "Sari":
        return 1
    return 0


def _build_platform_churn(series, max_level, reliable_from):
    """Platform churn tablolari: oncelikli leveller + bant ozeti."""
    start = series.get("Start", {})
    complete = series.get("Complete", {})
    fail = series.get("Fail", {})
    tryagain = series.get("Tryagain", {})
    if not complete:
        return {"hotspots": [], "bands": []}

    lo = max(1, int(reliable_from))
    hi = min(max_level, 200)
    rows = []
    for lvl in range(lo, hi + 1):
        s = start.get(lvl, 0.0)
        c = complete.get(lvl, 0.0)
        k = _excel_k_pct(complete, lvl)
        d1 = _drop1_pct(complete, lvl)
        k_cls = _classify_k(lvl, k)
        d_cls = _classify_drop1(d1)
        prio = _churn_priority_score(k_cls, d_cls)
        c_rate = (c / s * 100.0) if s > 0 else None
        note = []
        if d_cls == "Kirmizi":
            note.append("Ani dusus (H)")
        if k_cls == "Kirmizi":
            note.append("Dusuk retention (K)")
        if fail.get(lvl, 0) > 0 and c > 0 and fail.get(lvl, 0) / c > 0.3:
            note.append("Yuksek Fail")
        rows.append(
            {
                "level": lvl,
                "start": int(s),
                "complete": int(c),
                "complete_rate": round(c_rate, 1) if c_rate is not None else None,
                "k_pct": round(k, 1) if k is not None else None,
                "k_class": k_cls,
                "drop1_pct": round(d1, 1) if d1 is not None else None,
                "drop1_class": d_cls,
                "fail": int(fail.get(lvl, 0)),
                "tryagain": int(tryagain.get(lvl, 0)),
                "priority": prio,
                "note": "; ".join(note),
            }
        )

    hotspots = sorted(
        [r for r in rows if r["priority"] > 0],
        key=lambda x: (-x["priority"], -(x["drop1_pct"] or 0), -(x["k_pct"] or 0)),
    )[:30]

    band_defs = (
        ("Erken", lo, min(30, hi)),
        ("Orta", 31, min(100, hi)),
        ("Gec", 101, hi),
    )
    bands = []
    for name, b_lo, b_hi in band_defs:
        if b_lo > b_hi:
            continue
        band_rows = [r for r in rows if b_lo <= r["level"] <= b_hi]
        if not band_rows:
            continue
        rates = [r["complete_rate"] for r in band_rows if r["complete_rate"] is not None]
        drops = [r["drop1_pct"] for r in band_rows if r["drop1_pct"] is not None]
        bands.append(
            {
                "name": name,
                "range": f"L{b_lo}-L{b_hi}",
                "avg_complete_rate": round(sum(rates) / len(rates), 1) if rates else None,
                "avg_drop1": round(sum(drops) / len(drops), 1) if drops else None,
                "red_drop": sum(1 for r in band_rows if r["drop1_class"] == "Kirmizi"),
                "yellow_drop": sum(1 for r in band_rows if r["drop1_class"] == "Sari"),
                "red_k": sum(1 for r in band_rows if r["k_class"] == "Kirmizi"),
                "levels_in_band": len(band_rows),
            }
        )

    return {"hotspots": hotspots, "bands": bands}


def build_data_quality_report(
    data_android,
    data_ios,
    *,
    game_name,
    date_range,
    country_code,
    max_level,
    source_files,
):
    """Tum platformlar icin ozet + kalite + churn raporu."""
    platforms = {}
    for label, data in (("Android", data_android), ("iOS", data_ios)):
        if not any(
            data.get(ft) is not None and not data.get(ft).empty for ft in ("Start", "Complete")
        ):
            platforms[label] = {
                "label": label,
                "recommended_level": 1,
                "recommended_funnel_level": 1,
                "recommended_ads_level": 1,
                "funnel_first_level": {},
                "warnings": ["Bu platform icin Start/Complete verisi yok."],
                "level_rows": [],
                "series": {},
                "churn": {"hotspots": [], "bands": []},
            }
            continue
        plat = _analyze_platform_quality(data, max_level, label)
        plat["churn"] = _build_platform_churn(
            plat["series"],
            max_level,
            plat.get("recommended_funnel_level", 1),
        )
        platforms[label] = plat

    global_warn = [
        "Analytics'te user property ile segmentasyon yapildiginda, ozellik "
        "erken level'larda kullanici sayisi dusuk veya 0 gorunebilir. Bu gercek "
        "drop degil; ozellik henuz set edilmemis veya funnel filtresine girmemis "
        "kullanicilardan kaynaklanir.",
        "And/IOS/K (%) analizinde Level 1-2 yerine asagidaki 'Onerilen baslangic' "
        "level'inden itibaren yorum yapin.",
        "Drop-off (K) ve Drop 1 (H) formulleri dusuk level'de yanıltici olabilir; "
        "guvenilir band sonrasi trendlere odaklanin.",
    ]
    for pk, short in (("Android", "And"), ("iOS", "IOS")):
        p = platforms.get(pk, {})
        rf = p.get("recommended_funnel_level", p.get("recommended_level", 1))
        ra = p.get("recommended_ads_level", 1)
        global_warn.append(
            f"{pk}: funnel (K/H) L{rf}+ | ADs/rewarded (T) L{ra}+"
        )

    return {
        "game_name": game_name,
        "date_range": date_range,
        "country_code": country_code,
        "max_level": max_level,
        "source_files": [os.path.basename(p) for p in source_files],
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "platforms": platforms,
        "global_warnings": global_warn,
    }


def _ozet_set_cell(ws, row, col, value, *, bold=False, fill=None, wrap=False):
    c = ws.cell(row=row, column=col, value=value)
    if bold:
        c.font = Font(bold=True)
    if fill:
        c.fill = fill
    if wrap:
        c.alignment = Alignment(wrap_text=True, vertical="top")
    return c


def _ozet_row_fill(k_cls, d_cls, warn_fill, bad_fill):
    if d_cls == "Kirmizi" or k_cls == "Kirmizi":
        return bad_fill
    if d_cls == "Sari" or k_cls == "Sari":
        return warn_fill
    return None


def _write_ozet_table(ws, start_row, title, headers, rows, section_fill, warn_fill, bad_fill):
    """Baslik + tablo; satir sonunda ('__class__', k_cls, d_cls) ile renklendirme."""
    r = start_row
    _ozet_set_cell(ws, r, 1, title, bold=True, fill=section_fill)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=max(len(headers), 2))
    r += 1
    for ci, h in enumerate(headers, 1):
        _ozet_set_cell(ws, r, ci, h, bold=True, fill=section_fill)
    r += 1
    for row in rows:
        k_cls = d_cls = None
        if row and isinstance(row[-1], tuple) and row[-1][0] == "__class__":
            k_cls, d_cls = row[-1][1], row[-1][2]
            row = row[:-1]
        fill = _ozet_row_fill(k_cls, d_cls, warn_fill, bad_fill) if (k_cls or d_cls) else None
        for ci, val in enumerate(row, 1):
            _ozet_set_cell(ws, r, ci, val, fill=fill)
        r += 1
    return r + 1


def write_ozet_sheet(wb, report):
    """Excel'in ilk sekmesi: tablo halinde genel bilgi + churn ozeti."""
    if SUMMARY_SHEET_NAME in wb.sheetnames:
        del wb[SUMMARY_SHEET_NAME]
    ws = wb.create_sheet(SUMMARY_SHEET_NAME, 0)

    section_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    warn_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    bad_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    r = 1
    ws.cell(row=r, column=1, value="Ozet — Veri kalitesi ve churn").font = Font(bold=True, size=14)
    r += 2

    r = _write_ozet_table(
        ws,
        r,
        "1) Genel bilgi",
        ["Alan", "Deger"],
        [
            ["Oyun", report["game_name"]],
            ["Ulke", report["country_code"]],
            ["Tarih araligi", report["date_range"]],
            ["Max level", report["max_level"]],
            ["Olusturma", report["generated_at"]],
            ["Kaynak dosya", len(report["source_files"])],
        ],
        section_fill,
        warn_fill,
        bad_fill,
    )

    plat_rows = []
    for pk in ("Android", "iOS"):
        p = report["platforms"].get(pk, {})
        plat_rows.append(
            [
                pk,
                f"L{p.get('recommended_funnel_level', 1)}+",
                f"L{p.get('recommended_ads_level', 1)}+",
                len(p.get("churn", {}).get("hotspots", [])),
            ]
        )
    r = _write_ozet_table(
        ws,
        r,
        "2) Platform — guvenilir analiz bandi",
        ["Platform", "Funnel (K,H)", "ADs (T)", "Churn uyarisi sayisi"],
        plat_rows,
        section_fill,
        warn_fill,
        bad_fill,
    )

    r = _write_ozet_table(
        ws,
        r,
        "3) Notlar (segmentasyon / erken level)",
        ["Aciklama"],
        [[line] for line in report["global_warnings"][:6]],
        section_fill,
        warn_fill,
        bad_fill,
    )

    churn_rows = []
    for pk in ("Android", "iOS"):
        for item in report["platforms"].get(pk, {}).get("churn", {}).get("hotspots", []):
            churn_rows.append(
                [
                    pk,
                    item["level"],
                    item["start"],
                    item["complete"],
                    item["complete_rate"] if item["complete_rate"] is not None else "",
                    item["k_pct"] if item["k_pct"] is not None else "",
                    item["k_class"],
                    item["drop1_pct"] if item["drop1_pct"] is not None else "",
                    item["drop1_class"],
                    item["fail"],
                    item["tryagain"],
                    item["note"],
                    ("__class__", item["k_class"], item["drop1_class"]),
                ]
            )
    if not churn_rows:
        churn_rows = [["—", "—", "", "", "", "", "", "", "", "", "", "Uyari yok"]]
    r = _write_ozet_table(
        ws,
        r,
        "4) Churn — oncelikli leveller (Kirmizi/Sari K veya H)",
        [
            "Platform",
            "Level",
            "Start",
            "Complete",
            "Complete/Start %",
            "K %",
            "K durum",
            "Drop1 %",
            "H durum",
            "Fail",
            "Tryagain",
            "Not",
        ],
        churn_rows,
        section_fill,
        warn_fill,
        bad_fill,
    )

    band_rows = []
    for pk in ("Android", "iOS"):
        for b in report["platforms"].get(pk, {}).get("churn", {}).get("bands", []):
            band_rows.append(
                [
                    pk,
                    b["name"],
                    b["range"],
                    b["avg_complete_rate"] if b["avg_complete_rate"] is not None else "",
                    b["avg_drop1"] if b["avg_drop1"] is not None else "",
                    b["red_drop"],
                    b["yellow_drop"],
                    b["red_k"],
                    b["levels_in_band"],
                ]
            )
    if band_rows:
        r = _write_ozet_table(
            ws,
            r,
            "5) Level banti ozeti",
            [
                "Platform",
                "Bant",
                "Aralik",
                "Ort Complete/Start %",
                "Ort Drop1 %",
                "Kirmizi H",
                "Sari H",
                "Kirmizi K",
                "Level sayisi",
            ],
            band_rows,
            section_fill,
            warn_fill,
            bad_fill,
        )

    _write_ozet_table(
        ws,
        r,
        "6) Kaynak dosyalar",
        ["Dosya"],
        [[fn] for fn in report["source_files"]],
        section_fill,
        warn_fill,
        bad_fill,
    )

    widths = [14, 8, 10, 10, 14, 8, 10, 10, 10, 8, 10, 28]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    print(f"   OK '{SUMMARY_SHEET_NAME}' sekmesi eklendi (tablo ozeti + churn).")


def print_quality_report_console(report):
    """Konsola kisa ozet."""
    print("\n--- Ozet ---")
    print(f"  Oyun: {report['game_name']} | Tarih: {report['date_range']}")
    for pk in ("Android", "iOS"):
        p = report["platforms"].get(pk, {})
        rf = p.get("recommended_funnel_level", p.get("recommended_level", 1))
        ra = p.get("recommended_ads_level", 1)
        hot = p.get("churn", {}).get("hotspots", [])
        top = ", ".join(f"L{x['level']}" for x in hot[:5]) if hot else "yok"
        print(f"  {pk}: funnel L{rf}+ | ADs L{ra}+ | churn uyarisi: {top}")
    print("  Detay tablolar: Excel > 'Ozet' sekmesi.\n")


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

    quality_report = build_data_quality_report(
        data_android,
        data_ios,
        game_name=detected_game,
        date_range=date_range,
        country_code=country_code,
        max_level=max_level,
        source_files=files,
    )
    write_ozet_sheet(wb, quality_report)
    print_quality_report_console(quality_report)

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
