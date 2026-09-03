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


def _configure_stdout_utf8():
    """Windows terminalinde emoji/unicode dosya adlarini yazdirabilmek icin."""
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    if hasattr(sys.stderr, "reconfigure"):
        try:
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def _safe_console_text(text):
    s = str(text)
    enc = getattr(sys.stdout, "encoding", None) or "utf-8"
    try:
        s.encode(enc)
        return s
    except (UnicodeEncodeError, LookupError):
        return s.encode(enc, errors="replace").decode(enc, errors="replace")


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


def _is_revenue_level_table_file(basename: str) -> bool:
    """Analytics detailed level table (Level + Revenue); funnel verisi degil."""
    return "detailed_level_table" in str(basename).lower()


def _is_revenue_level_table_xlsx(basename: str) -> bool:
    return _is_revenue_level_table_file(basename) and str(basename).lower().endswith(
        (".xlsx", ".xls", ".xlsm")
    )


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
    """# hole-pool etiketi veya Free form tarzi (OS/Platform veya Level+funnel) yapi."""
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
    if "level" not in blob:
        return False
    if not any(k in blob for k in ("start", "complete", "fail", "tryagain", "ads")):
        return False
    if "operating system" in blob or "platform" in blob:
        return True
    # Platform yok: App version / network kirilimli export
    return any(
        k in blob
        for k in ("app version", "user_network", "user network")
    )


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
        if _is_revenue_level_table_file(base):
            skipped.append((base, "revenue tablosu"))
            continue
        try:
            df_check = pd.read_excel(path, nrows=50, header=None)
        except Exception as e:
            skipped.append((base, f"okunamadi: {e!s}"))
            continue
        if not _xlsx_rows_look_like_data(df_check):
            skipped.append((base, "veri yapisi (oyun etiketi / Level+funnel) yok"))
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


def get_unspecified_platform_target_from_user(file_names):
    """
    Export'ta Platform/OS yoksa And / IOS / her ikisi secimi.
    Enter veya EOF: And (varsayilan).
    """
    names = [str(n) for n in (file_names or []) if n]
    print(
        f"   UYARI: {len(names)} dosyada Platform (Android/iOS) yok "
        "(App version / network kirilimi)."
    )
    if names:
        shown = names[:8]
        extra = f" (+{len(names) - 8})" if len(names) > 8 else ""
        print("          " + ", ".join(shown) + extra)
    prompt = (
        "   Veri nereye yazilsin? And / IOS / both  (Enter = And) : "
    )
    while True:
        try:
            raw = input(prompt).strip().lower()
        except EOFError:
            return "and"
        if not raw or raw in ("and", "android", "1"):
            return "and"
        if raw in ("ios", "i", "2"):
            return "ios"
        if raw in ("both", "her ikisi", "heriki", "all", "3"):
            return "both"
        print("   Gecerli secim: And, IOS veya both.")


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

# Platform disi kirilim kolonlari (metrik degil; Level bazinda toplanir)
_DIMENSION_COL_CANON = frozenset(
    {
        "event_name",
        "eventname",
        "app_version",
        "appversion",
        "user_network",
        "usernetwork",
        "network",
        "segment",
    }
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


def _is_os_header_label(value):
    """Analytics export: 'Operating system' (ADs) veya 'Platform' (Start/Complete vb.)."""
    return str(value).strip().casefold() in ("operating system", "platform")


def _platform_column_name(columns):
    for name in ("Operating system", "Platform"):
        if name in columns:
            return name
    return None


def _header_cell_name(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def _unique_header_names(names):
    seen = {}
    unique = []
    for c in names:
        if c in seen:
            seen[c] += 1
            unique.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            unique.append(c)
    return unique


def _is_blank_header(name):
    s = str(name).strip()
    if not s:
        return True
    base = str(_base_col_name(s)).strip()
    return not base


def _is_dimension_column(name):
    s = str(name)
    k = _norm_metric_token(_base_col_name(s))
    k2 = _norm_metric_token(s)
    return k in _DIMENSION_COL_CANON or k2 in _DIMENSION_COL_CANON


def _drop_non_metric_columns(df):
    """Event name / App version / user_network ve bos basliklari at."""
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()
    drop = [
        c
        for c in df.columns
        if c != "Level" and (_is_blank_header(c) or _is_dimension_column(c))
    ]
    if not drop:
        return df
    return df.drop(columns=drop, errors="ignore")


def _keep_numeric_levels(df):
    if df is None or df.empty or "Level" not in df.columns:
        return df if df is not None else pd.DataFrame()
    out = df.copy()
    lvl = pd.to_numeric(out["Level"], errors="coerce")
    return out.loc[lvl.notna()].reset_index(drop=True)


def _dataframe_from_header_row(df_raw, header_idx):
    header_vals = [_header_cell_name(x) for x in df_raw.iloc[header_idx].tolist()]
    unique_cols = _unique_header_names(header_vals)
    df_flat = df_raw.iloc[header_idx + 1 :].copy()
    df_flat.columns = unique_cols
    if "Level" not in df_flat.columns:
        return pd.DataFrame()
    return _keep_numeric_levels(df_flat)


def read_platform_data(f):
    """
    Excel dosyasini okur, Android ve iOS verilerini ayri DataFrame olarak dondurur.
    Platform yoksa (App version / network kirilimi) tum satirlar df_android'da,
    has_platform=False doner.
    Doner: (funnel_type, df_android, df_ios, has_platform)
    """
    df_raw = pd.read_excel(f, header=None)
    funnel_type = detect_funnel_type(df_raw)
    empty = (funnel_type, pd.DataFrame(), pd.DataFrame(), False)

    os_row_idx = None
    level_row_idx = None
    for idx in range(min(15, len(df_raw))):
        row = df_raw.iloc[idx]
        for cell in row:
            if _is_os_header_label(cell) and os_row_idx is None:
                os_row_idx = idx
            if _header_cell_name(cell) == "Level" and level_row_idx is None:
                level_row_idx = idx

    if level_row_idx is None:
        print(f"   UYARI: '{f}' -> 'Level' bulunamadi, atlaniyor.")
        return empty

    if os_row_idx is None:
        df_flat = _drop_non_metric_columns(
            _dataframe_from_header_row(df_raw, level_row_idx)
        )
        print(
            f"   Bilgi: '{os.path.basename(f)}' -> Platform yok; "
            "App version / network satirlari Level bazinda toplanacak."
        )
        return funnel_type, df_flat, pd.DataFrame(), False

    os_row_vals = df_raw.iloc[os_row_idx].fillna("").astype(str).str.strip().tolist()
    level_row_vals = [_header_cell_name(x) for x in df_raw.iloc[level_row_idx].tolist()]

    if level_row_idx == os_row_idx:
        df_flat = _dataframe_from_header_row(df_raw, level_row_idx)
        os_col = _platform_column_name(df_flat.columns)
        if os_col is None:
            df_flat = _drop_non_metric_columns(df_flat)
            print(
                f"   Bilgi: '{os.path.basename(f)}' -> Platform sutunu yok; "
                "Level bazinda toplanacak."
            )
            return funnel_type, df_flat, pd.DataFrame(), False

        df_and = df_flat[
            df_flat[os_col].astype(str).str.contains("Android", case=False, na=False)
        ].copy()
        df_ios = df_flat[
            df_flat[os_col].astype(str).str.contains("iOS", case=False, na=False)
        ].copy()

        df_and = _drop_non_metric_columns(
            df_and.drop(columns=[os_col], errors="ignore").reset_index(drop=True)
        )
        df_ios = _drop_non_metric_columns(
            df_ios.drop(columns=[os_col], errors="ignore").reset_index(drop=True)
        )
        return funnel_type, df_and, df_ios, True

    level_col_idx = next((i for i, v in enumerate(level_row_vals) if v == "Level"), None)
    if level_col_idx is None:
        print(f"   UYARI: '{f}' -> Level sutun indeksi bulunamadi.")
        return empty

    and_indices = [i for i, v in enumerate(os_row_vals) if v == "Android"]
    ios_indices = [i for i, v in enumerate(os_row_vals) if v == "iOS"]
    data_start = level_row_idx + 1

    def extract(col_indices):
        cols = [level_col_idx] + col_indices
        sub = df_raw.iloc[data_start:, cols].copy()
        raw_names = [level_row_vals[level_col_idx]] + [
            level_row_vals[i] for i in col_indices
        ]
        sub.columns = _unique_header_names(raw_names)
        return _drop_non_metric_columns(_keep_numeric_levels(sub))

    return funnel_type, extract(and_indices), extract(ios_indices), True


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
    if df_sub.empty or "Level" not in df_sub.columns:
        return pd.DataFrame({"Level": range(1, max_level + 1)})

    df_sub = _drop_non_metric_columns(df_sub.copy())
    df_sub["Level"] = pd.to_numeric(df_sub["Level"], errors="coerce")
    df_sub = df_sub.dropna(subset=["Level"]).copy()
    df_sub["Level"] = df_sub["Level"].astype(int)
    df_sub = df_sub[df_sub["Level"] <= max_level]

    for col in df_sub.columns:
        if col != "Level":
            df_sub[col] = pd.to_numeric(df_sub[col], errors="coerce").fillna(0)

    numeric_cols = df_sub.select_dtypes(include="number").columns.tolist()
    if "Level" in numeric_cols:
        numeric_cols.remove("Level")

    all_levels = pd.DataFrame({"Level": range(1, max_level + 1)})
    if not numeric_cols:
        return all_levels

    grouped = df_sub.groupby("Level")[numeric_cols].sum().reset_index()
    grouped = grouped.sort_values(by="Level")
    merged = pd.merge(all_levels, grouped, on="Level", how="left")
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


def _header_cell_display(value, fallback=""):
    if value is None:
        return fallback
    s = str(value).strip()
    return s if s else fallback


# And/IOS churn metrikleri: H (8) ve K (11) — satir 2 alt basliklari
_AND_IOS_K_COL = 11
_AND_IOS_H_COL = 8
_AND_IOS_METRIC_HEADER_ROW = 2


def _extract_and_ios_metric_labels(wb):
    """And/IOS 2. satirdaki gercek sutun adlarini okur (or. % ve Drop 1)."""
    ws = _worksheet_by_name(wb, "And") or _worksheet_by_name(wb, "IOS")
    if ws is None:
        return {"k_label": "%", "h_label": "Drop 1"}
    r = _AND_IOS_METRIC_HEADER_ROW
    return {
        "k_label": _header_cell_display(
            ws.cell(row=r, column=_AND_IOS_K_COL).value, "%"
        ),
        "h_label": _header_cell_display(
            ws.cell(row=r, column=_AND_IOS_H_COL).value, "Drop 1"
        ),
    }


def _report_metric_labels(report):
    labels = report.get("metric_labels") or {}
    return {
        "k_label": labels.get("k_label") or "%",
        "h_label": labels.get("h_label") or "Drop 1",
    }


def _format_churn_hotspot_label(item, metric_labels=None):
    """Ornek: L10 (%=Yesil, Drop 1=Sari) — And/IOS satir 2 basliklariyla."""
    labels = metric_labels or {}
    k_name = labels.get("k_label") or "%"
    h_name = labels.get("h_label") or "Drop 1"
    return (
        f"L{item['level']} "
        f"({k_name}={item.get('k_class', '')}, {h_name}={item.get('drop1_class', '')})"
    )


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


def _resolve_detected_game_name(game_dict):
    """Metaveriden en sik gecen oyun adini sec."""
    game_names = [g for g in (game_dict or {}).values() if g]
    if not game_names:
        return None
    game_counter = {}
    for g in game_names:
        game_counter[g] = game_counter.get(g, 0) + 1
    return sorted(game_counter.items(), key=lambda x: (-x[1], x[0]))[0][0]


def find_arpu_reference_workbook(script_dir, output_path=None):
    """
    Oncelik: ARPU_CHART_REFERENCE_XLSX. Yoksa en yeni *_Otomatik_Analiz*.xlsx
    (cikti dosyasi ve sablon haric).
    """
    out_abs = os.path.abspath(output_path) if output_path else None
    pref = os.path.join(script_dir, ARPU_CHART_REFERENCE_XLSX)
    if os.path.isfile(pref) and (out_abs is None or os.path.abspath(pref) != out_abs):
        return pref
    paths = glob.glob(os.path.join(script_dir, "*_Otomatik_Analiz*.xlsx"))
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
    # - Level 1 (satir 3): BOŞ — Start user property sebebiyle L1 temiz gelmiyor;
    #   C3 bazli hesap yaniltici olur.
    # - Level 2..100: K(n) = 100-(((C(n-1)-C(n))/C(n-1))*100)
    #   Yani her satir bir onceki level'in Complete'ini baz alir.
    # - Level >100 : boş bırakılır (gösterilmez)
    if sheet_name in ('And', 'IOS'):
        l1_row = level_rows.get(1, 3)  # C3 = L1 Complete satiri
        for level in range(1, max_level + 1):
            t_row = level_rows[level]
            if level == 1:
                # K3 bos; L1 baz satiri, kendisiyle karsilastirma anlamsiz
                ws.cell(row=t_row, column=11).value = None
            elif level <= 100:
                # K(n) = 100 - ((C3 - C(n)) / C3 * 100)  — hep L1 Complete baz
                ws.cell(row=t_row, column=11).value = (
                    f'=100-(((C{l1_row}-C{t_row})/C{l1_row})*100)'
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
    """
    And/IOS K (%) sutunu ile ayni mantik: hep L1 Complete (C3) baz.
    K(n) = 100 - ((C1 - C(n)) / C1 * 100)
    Level 1 icin None donulur (K3 bos).
    """
    if level <= 1:
        return None
    c1 = complete.get(1, 0.0)
    c_curr = complete.get(level, 0.0)
    if c1 <= 0:
        return None
    return 100.0 - ((c1 - c_curr) / c1 * 100.0)


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
        key=lambda x: x["level"],
    )[:50]

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


def _fmt_int(v):
    try:
        return f"{int(round(float(v))):,}".replace(",", ".")
    except Exception:
        return "0"


def _sum_series_values(level_map):
    if not isinstance(level_map, dict):
        return 0.0
    return float(sum(v for v in level_map.values() if pd.notna(v)))


def _build_revenue_summary(revenue_android_df, revenue_ios_df, revenue_source):
    def _sum_rev(df):
        if df is None or df.empty or "Total $" not in df.columns:
            return 0.0
        s = pd.to_numeric(df["Total $"], errors="coerce").fillna(0).sum()
        return float(s)

    and_total = _sum_rev(revenue_android_df)
    ios_total = _sum_rev(revenue_ios_df)
    total = and_total + ios_total
    if total <= 0:
        return {}
    and_share = (and_total / total * 100.0) if total > 0 else 0.0
    ios_share = (ios_total / total * 100.0) if total > 0 else 0.0
    return {
        "source": revenue_source or "",
        "android_total": and_total,
        "ios_total": ios_total,
        "total": total,
        "android_share_pct": and_share,
        "ios_share_pct": ios_share,
    }


def _build_auto_insights(report):
    """
    Esnek insight listesi:
    - Sadece raporda gercekten bulunan verilerden cikarim yazar.
    - Veri yoksa ilgili satir eklenmez (hallucination yok).
    """
    insights = []
    game = report.get("game_name", "Oyun")
    date_range = report.get("date_range", "-")
    max_level = report.get("max_level", 0)
    insights.append(f"{game} | {date_range} donemi, analiz limiti L{max_level}.")

    rev = report.get("revenue_summary") or {}
    if rev:
        insights.append(
            "Purchase revenue toplam $"
            f"{rev.get('total', 0):,.0f} "
            f"(Android ${rev.get('android_total', 0):,.0f}, iOS ${rev.get('ios_total', 0):,.0f}; "
            f"pay: %{rev.get('android_share_pct', 0):.1f} / %{rev.get('ios_share_pct', 0):.1f})."
        )
        if rev.get("source"):
            insights.append(f"Gelir kaynagi: {rev['source']}.")

    for pk in ("Android", "iOS"):
        p = report["platforms"].get(pk, {})
        series = p.get("series", {})
        start_l1 = float(series.get("Start", {}).get(1, 0.0))
        complete_l1 = float(series.get("Complete", {}).get(1, 0.0))
        fail_total = _sum_series_values(series.get("Fail", {}))
        try_total = _sum_series_values(series.get("Tryagain", {}))
        ads_total = _sum_series_values(series.get("ADs", {}))
        rf = p.get("recommended_funnel_level", p.get("recommended_level", 1))
        ra = p.get("recommended_ads_level", 1)

        if start_l1 > 0 or complete_l1 > 0:
            insights.append(
                f"{pk}: L1 Start {_fmt_int(start_l1)}, L1 Complete {_fmt_int(complete_l1)} | "
                f"guvenilir funnel bandi L{rf}+."
            )
        if ads_total > 0:
            insights.append(
                f"{pk}: toplam rewarded/ad event {_fmt_int(ads_total)} "
                f"(ADs yorum bandi L{ra}+)."
            )
        if fail_total > 0 or try_total > 0:
            insights.append(
                f"{pk}: toplam Fail {_fmt_int(fail_total)} | Tryagain {_fmt_int(try_total)}."
            )

        hot = p.get("churn", {}).get("hotspots", [])
        if hot:
            top = hot[:3]
            labels = _report_metric_labels(report)
            hotspot_txt = ", ".join(
                _format_churn_hotspot_label(x, labels) for x in top
            )
            insights.append(f"{pk}: oncelikli churn levelleri -> {hotspot_txt}.")

    if not insights:
        insights = ["Insight uretilemedi (yetersiz veri)."]
    return insights


def _extract_platform_metrics_from_sheet(wb, sheet_name, max_level):
    """
    Insight icin dogrudan otomatik analiz workbook tablosundan metrik ceker.
    And/IOS kolon semasi:
      B: Start, C: Complete, D: Fail, E: Tryagain, N: Total $, T: rewarded ratio
    """
    ws = _worksheet_by_name(wb, sheet_name)
    if ws is None:
        return {}
    level_rows = find_level_rows(ws)
    if not level_rows:
        return {}

    def _cell_num(r, c):
        try:
            return float(pd.to_numeric(ws.cell(row=r, column=c).value, errors="coerce"))
        except Exception:
            return 0.0

    out = {
        "start_total": 0.0,
        "complete_total": 0.0,
        "fail_total": 0.0,
        "tryagain_total": 0.0,
        "revenue_total": 0.0,
        "rewarded_ratio_avg": 0.0,
        "rewarded_ratio_count": 0,
        "l1_start": 0.0,
        "l1_complete": 0.0,
    }
    for lvl in range(1, max_level + 1):
        r = level_rows.get(lvl)
        if not r:
            continue
        s = _cell_num(r, 2)
        c = _cell_num(r, 3)
        f = _cell_num(r, 4)
        t = _cell_num(r, 5)
        rev = _cell_num(r, 14)
        rew_ratio = _cell_num(r, 20)

        if lvl == 1:
            out["l1_start"] = s
            out["l1_complete"] = c
        out["start_total"] += s
        out["complete_total"] += c
        out["fail_total"] += f
        out["tryagain_total"] += t
        out["revenue_total"] += rev
        if rew_ratio > 0:
            out["rewarded_ratio_avg"] += rew_ratio
            out["rewarded_ratio_count"] += 1

    if out["rewarded_ratio_count"] > 0:
        out["rewarded_ratio_avg"] = out["rewarded_ratio_avg"] / out["rewarded_ratio_count"]
    return out


def _extract_workbook_insight_context(wb, max_level):
    return {
        "Android": _extract_platform_metrics_from_sheet(wb, "And", max_level),
        "iOS": _extract_platform_metrics_from_sheet(wb, "IOS", max_level),
    }


def _apply_workbook_context_to_report(report, workbook_context):
    """
    Insight dilini workbook metriklerine yaslamak icin rapora tablo-odakli ozet ekler.
    """
    if not workbook_context:
        return report
    r = dict(report)
    r["workbook_context"] = workbook_context
    return r


def _build_auto_insights_from_workbook(report):
    """
    Oncelik: otomatik analiz dosyasina yazilan And/IOS tablo degerleri.
    """
    ctx = report.get("workbook_context") or {}
    if not ctx:
        return _build_auto_insights(report)

    insights = []
    game = report.get("game_name", "Oyun")
    date_range = report.get("date_range", "-")
    max_level = report.get("max_level", 0)
    insights.append(f"{game} | {date_range} donemi, dosya bazli analiz limiti L{max_level}.")

    and_m = ctx.get("Android", {})
    ios_m = ctx.get("iOS", {})
    total_rev = float(and_m.get("revenue_total", 0.0)) + float(ios_m.get("revenue_total", 0.0))
    if total_rev > 0:
        a_rev = float(and_m.get("revenue_total", 0.0))
        i_rev = float(ios_m.get("revenue_total", 0.0))
        insights.append(
            f"Dosyaya yazilan toplam purchase revenue ${total_rev:,.0f} "
            f"(Android ${a_rev:,.0f}, iOS ${i_rev:,.0f})."
        )

    for pk, m in (("Android", and_m), ("iOS", ios_m)):
        if not m:
            continue
        l1s = float(m.get("l1_start", 0.0))
        l1c = float(m.get("l1_complete", 0.0))
        if l1s > 0 or l1c > 0:
            insights.append(
                f"{pk}: dosya tablosunda L1 Start {_fmt_int(l1s)}, L1 Complete {_fmt_int(l1c)}."
            )
        f_total = float(m.get("fail_total", 0.0))
        t_total = float(m.get("tryagain_total", 0.0))
        if f_total > 0 or t_total > 0:
            insights.append(
                f"{pk}: toplam Fail {_fmt_int(f_total)} | Tryagain {_fmt_int(t_total)} (dosya toplam)."
            )
        rew_avg = float(m.get("rewarded_ratio_avg", 0.0))
        if rew_avg > 0:
            insights.append(
                f"{pk}: ortalama rewarded/complete orani ~{rew_avg:.2f} (T sutunu ortalamasi)."
            )

    # Churn ve kalite notlari yine kalite raporundan gelir
    for pk in ("Android", "iOS"):
        p = report["platforms"].get(pk, {})
        hot = sorted(p.get("churn", {}).get("hotspots", []), key=lambda x: x["level"])
        if hot:
            top = hot[:5]
            labels = _report_metric_labels(report)
            hotspot_txt = ", ".join(
                _format_churn_hotspot_label(x, labels) for x in top
            )
            insights.append(f"{pk}: churn uyari levelleri (level sirasi) -> {hotspot_txt}.")

    return insights


def build_data_quality_report(
    data_android,
    data_ios,
    *,
    game_name,
    date_range,
    country_code,
    max_level,
    source_files,
    revenue_android_df=None,
    revenue_ios_df=None,
    revenue_source="",
    metric_labels=None,
):
    """Tum platformlar icin ozet + kalite + churn raporu."""
    metric_labels = metric_labels or {"k_label": "%", "h_label": "Drop 1"}
    k_label = metric_labels.get("k_label") or "%"
    h_label = metric_labels.get("h_label") or "Drop 1"
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
        f"And/IOS/{k_label} analizinde Level 1-2 yerine asagidaki 'Onerilen baslangic' "
        "level'inden itibaren yorum yapin.",
        f"{k_label} ve {h_label} formulleri dusuk level'de yanıltici olabilir; "
        "guvenilir band sonrasi trendlere odaklanin.",
    ]
    for pk, short in (("Android", "And"), ("iOS", "IOS")):
        p = platforms.get(pk, {})
        rf = p.get("recommended_funnel_level", p.get("recommended_level", 1))
        ra = p.get("recommended_ads_level", 1)
        global_warn.append(
            f"{pk}: funnel ({k_label}/{h_label}) L{rf}+ | ADs/rewarded (T) L{ra}+"
        )

    revenue_summary = _build_revenue_summary(
        revenue_android_df, revenue_ios_df, revenue_source
    )
    return {
        "game_name": game_name,
        "date_range": date_range,
        "country_code": country_code,
        "max_level": max_level,
        "source_files": [os.path.basename(p) for p in source_files],
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "platforms": platforms,
        "revenue_summary": revenue_summary,
        "global_warnings": global_warn,
        "metric_labels": metric_labels,
        "insights": [],
    }


def _ozet_metric_legend_rows(metric_labels=None):
    """And/IOS satir 2 basliklariyla Ozet metrik sozlugu."""
    labels = metric_labels or {}
    k_label = labels.get("k_label") or "%"
    h_label = labels.get("h_label") or "Drop 1"
    return [
        [
            k_label,
            f"And/IOS {k_label} sutunu (retention). L1 Complete'a gore bu levelda kalan oyuncu orani. "
            "Yuksek = iyi.",
        ],
        [
            f"{k_label} durumu",
            f"{k_label} renk esigi (level bandina gore): Yesil=iyi, Sari=orta, Kirmizi=dusuk retention.",
        ],
        [
            h_label,
            f"And/IOS {h_label} sutunu. Bu leveldan bir sonraki levela geciste Complete dususu (%). "
            "Ani kayip gostergesi.",
        ],
        [
            f"{h_label} durumu",
            f"{h_label} renk esigi: OK=<%8.5, Sari=%8.5-10, Kirmizi=>%10 (ani dusus).",
        ],
        [
            "Churn tablosu",
            f"Asagidaki leveller yalnizca Kirmizi/Sari {k_label} veya {h_label} olan satirlardir; "
            "level numarasina gore kucukten buyuge siralanir.",
        ],
        [
            "Guvenilir band",
            "Erken level segmentasyon artefakti olabilir; 'Platform — guvenilir analiz bandi' "
            "satirindaki Lx+ sonrasina odaklanin.",
        ],
    ]


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
            c = _ozet_set_cell(ws, r, ci, val, fill=fill)
            if ci == 1 and title.startswith("4) Metrik"):
                c.alignment = Alignment(wrap_text=True, vertical="top")
            if ci == 2 and title.startswith("4) Metrik"):
                c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    return r + 1


def write_ozet_sheet(wb, report):
    """Excel'in ilk sekmesi: tablo halinde genel bilgi + churn ozeti."""
    if SUMMARY_SHEET_NAME in wb.sheetnames:
        del wb[SUMMARY_SHEET_NAME]
    ws = wb.create_sheet(SUMMARY_SHEET_NAME, 0)
    metric_labels = _report_metric_labels(report)
    k_label = metric_labels["k_label"]
    h_label = metric_labels["h_label"]

    section_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    warn_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    bad_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    r = 1
    ws.cell(row=r, column=1, value="Ozet — Veri kalitesi ve churn").font = Font(bold=True, size=14)
    r += 2

    insight_rows = [[line] for line in (report.get("insights") or [])]
    if insight_rows:
        r = _write_ozet_table(
            ws,
            r,
            "0) Insights (otomatik)",
            ["Insight"],
            insight_rows,
            section_fill,
            warn_fill,
            bad_fill,
        )

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
        [
            "Platform",
            f"Funnel ({k_label},{h_label})",
            "ADs (T)",
            f"{k_label}/{h_label} uyarili level sayisi",
        ],
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

    r = _write_ozet_table(
        ws,
        r,
        f"4) Metrik sozlugu — {k_label} ve {h_label} (And/IOS satir 2)",
        ["Metrik", "Ne anlama geliyor?"],
        _ozet_metric_legend_rows(metric_labels),
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
    churn_rows.sort(key=lambda row: (0 if row[0] == "Android" else 1, int(row[1])))
    if not churn_rows:
        churn_rows = [["—", "—", "", "", "", "", "", "", "", "", "", "Uyari yok"]]
    r = _write_ozet_table(
        ws,
        r,
        f"5) Churn — uyari levelleri (level sirasina gore; {k_label}/{h_label} Kirmizi veya Sari)",
        [
            "Platform",
            "Level",
            "Start",
            "Complete",
            "Complete/Start %",
            k_label,
            f"{k_label} durumu",
            h_label,
            f"{h_label} durumu",
            "Fail",
            "Tryagain",
            "Kisa not",
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
            "6) Level banti ozeti",
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
        "7) Kaynak dosyalar",
        ["Dosya"],
        [[fn] for fn in report["source_files"]],
        section_fill,
        warn_fill,
        bad_fill,
    )

    widths = [14, 10, 10, 10, 14, 14, 12, 12, 12, 8, 10, 28]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    print(f"   OK '{SUMMARY_SHEET_NAME}' sekmesi eklendi (tablo ozeti + churn).")


def _platform_has_funnel_data(data_by_funnel):
    for ft in ("Start", "Complete"):
        df = data_by_funnel.get(ft)
        if df is not None and not df.empty:
            return True
    return False


def _parse_level_value(v):
    """'23' / 23 / 'Lv. 23' -> int; okunamazsa None."""
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        if float(v) != float(v):  # NaN
            return None
        return int(v)
    s = str(v).strip()
    m = re.search(r"(\d+)", s)
    if not m:
        return None
    return int(m.group(1))


def _parse_money_value(v):
    """12.5 / '$31' / '$10.39' -> float."""
    if v is None:
        return 0.0
    try:
        if pd.isna(v):
            return 0.0
    except (TypeError, ValueError):
        pass
    if isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = (
        str(v)
        .strip()
        .replace("$", "")
        .replace(",", "")
        .replace("\u00a0", "")
        .strip()
    )
    n = pd.to_numeric(s, errors="coerce")
    return float(n) if pd.notna(n) else 0.0


def _game_keys_match(cell, detected):
    if not detected:
        return True
    a = re.sub(r"[^a-z0-9]+", "", str(cell).lower())
    b = re.sub(r"[^a-z0-9]+", "", str(detected).lower())
    if not a or not b:
        return False
    return a == b or a in b or b in a


def _find_revenue_value_column(df):
    """Level tablosunda revenue sutununu bul (Revenue ($), Revenue, vb.)."""
    for col in df.columns:
        key = re.sub(r"\s+", " ", str(col).strip().lower())
        if key in ("revenue ($)", "revenue", "total $", "usd price", "revenue usd"):
            return col
        if "revenue" in key and "$" in key:
            return col
    return None


def _detect_revenue_platform_from_filename(path):
    """
    Dosya adindan Google (And) / iOS ayirimi.
    Ornek: ...google..., ...android..., ...ios..., ...appstore...
    """
    base = os.path.basename(path).lower()
    and_keys = ("google", "android", "playstore", "play_store", "play-store", "_and_")
    ios_keys = ("ios", "appstore", "app_store", "app-store", "iphone", "_ios_")
    if any(k in base for k in and_keys):
        return "and"
    if any(k in base for k in ios_keys):
        return "ios"
    return None


def _purchase_platform_display_label(platform_code):
    if platform_code == "and":
        return "Google (And)"
    if platform_code == "ios":
        return "iOS"
    return "BELIRSIZ"


def _print_purchase_export_naming_requirement(unlabeled_paths=None):
    """
    Purchase export dosya adinda platform etiketi yoksa uyari ve analytics ekibi icin metin.
    """
    print("\n   *** UYARI: Purchase dosya adinda platform (google/ios) yok ***")
    print("   Export dosyalari su an yalnizca zaman damgasi ile geliyor; And/IOS ayrimi")
    print("   dosya adindan veya CSV icinden okunamiyor.")
    print("")
    print("   Cozum (tercih edilen): Export/indirme tarafinda dosya adina platform eklensin:")
    print("     google_detailed_level_table_<tarih>.csv")
    print("     ios_detailed_level_table_<tarih>.csv")
    print("")
    print("   Kabul edilen anahtar kelimeler (dosya adinda, buyuk/kucuk harf fark etmez):")
    print("     Google/Android -> And    |    iOS/AppStore -> IOS")
    if unlabeled_paths:
        print("")
        print("   Etiketsiz dosyalar:")
        for p in unlabeled_paths:
            print(f"     - {_safe_console_text(os.path.basename(p))}")
    print("")
    print("   --- Analytics / export ekibine iletilebilecek metin ---")
    print(
        "   detailed_level_table purchase export'larinda dosya adina platform eklenmeli.\n"
        "   Ornek: google_detailed_level_table_2026-06-30.csv ve\n"
        "          ios_detailed_level_table_2026-06-30.csv\n"
        "   CSV icinde Store/OS kolonu olmadigi icin dosya adi zorunlu ayrim kaynagi."
    )
    print("   --------------------------------------------------------\n")


def _warn_unlabeled_purchase_files(unlabeled_paths, and_parts, ios_parts):
    if not unlabeled_paths:
        return
    _print_purchase_export_naming_requirement(unlabeled_paths)
    if and_parts or ios_parts:
        print(
            "   Not: Etiketsiz dosyalar atlandi (etiketli dosyalar kullanildi).\n"
            "        Lutfen etiketsiz exportlari yeniden adlandirin veya analytics ekibinden\n"
            "        platformlu dosya adi isteyin.\n"
        )


def _read_revenue_level_table(path, game_name=None):
    """
    Level + Revenue -> Level, Total $.
    detailed_level_table (sayisal Level, Revenue ($)) ve
    top_levels_by_purchase_count (Lv. 23, $31, Game) desteklenir.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(path)
    else:
        df = pd.read_excel(path)
    if df is None or df.empty:
        return pd.DataFrame()
    level_col = None
    game_col = None
    for col in df.columns:
        key = str(col).strip().lower()
        if key == "level" and level_col is None:
            level_col = col
        if key == "game" and game_col is None:
            game_col = col
    if level_col is None:
        return pd.DataFrame()
    rev_col = _find_revenue_value_column(df)
    if rev_col is None:
        return pd.DataFrame()
    if game_col is not None and game_name:
        mask = df[game_col].map(lambda x: _game_keys_match(x, game_name))
        if mask.any():
            df = df.loc[mask].copy()
        else:
            print(
                f"   UYARI: '{os.path.basename(path)}' icinde oyun "
                f"'{game_name}' satiri yok; bu dosyada gelir atlandi."
            )
            return pd.DataFrame()
    sub = df[[level_col, rev_col]].copy()
    sub.columns = ["Level", "Total $"]
    sub["Level"] = sub["Level"].map(_parse_level_value)
    sub["Total $"] = sub["Total $"].map(_parse_money_value)
    sub = sub.dropna(subset=["Level"])
    sub["Level"] = sub["Level"].astype(int)
    out = sub.groupby("Level", as_index=False)["Total $"].sum()
    return out.sort_values("Level").reset_index(drop=True)


def _prompt_revenue_file_platforms(unlabeled_paths):
    """Dosya adinda platform yoksa kullaniciya sor (Google=And, diger=IOS)."""
    if len(unlabeled_paths) != 2:
        return {}
    _print_purchase_export_naming_requirement(unlabeled_paths)
    print("   Simdilik hangi dosyanin Google oldugunu secin.")
    print("   (Bir sonraki analiz icin dosyalari google_/ios_ ile yeniden adlandirin.)\n")
    for i, p in enumerate(unlabeled_paths, 1):
        print(f"     {i}) {_safe_console_text(os.path.basename(p))}")
    while True:
        try:
            raw = input(
                "   Google (And) purchase hangi dosya? (1 veya 2; Enter=1) : "
            ).strip()
            if not raw:
                idx_and = 0
            else:
                idx_and = int(raw) - 1
            if idx_and in (0, 1):
                idx_ios = 1 - idx_and
                return {
                    "and": unlabeled_paths[idx_and],
                    "ios": unlabeled_paths[idx_ios],
                }
        except (ValueError, EOFError):
            return {
                "and": unlabeled_paths[0],
                "ios": unlabeled_paths[1],
            }
        print("   1 veya 2 girin.")


def _merge_revenue_parts(parts):
    if not parts:
        return pd.DataFrame()
    merged = pd.concat(parts, ignore_index=True)
    merged = merged.groupby("Level", as_index=False)["Total $"].sum()
    return merged.sort_values("Level").reset_index(drop=True)


def _collect_revenue_level_table_paths(script_dir):
    """
    Oncelik: detailed_level_table.
    Yoksa: top_levels_by_purchase_count / purchase_count_game_breakdown.
    """
    detailed = []
    for pattern in ("*detailed_level_table*.csv", "*detailed_level_table*.xlsx"):
        detailed.extend(glob.glob(os.path.join(script_dir, pattern)))
    detailed = [
        p
        for p in detailed
        if not _is_output_workbook(os.path.basename(p))
        and not _is_template_workbook(os.path.basename(p))
    ]
    if detailed:
        detailed.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        return detailed

    breakdown = []
    for pattern in (
        "*top_levels_by_purchase*.csv",
        "*purchase_count_game_breakdown*.csv",
    ):
        breakdown.extend(glob.glob(os.path.join(script_dir, pattern)))
    seen = set()
    uniq = []
    for p in breakdown:
        key = os.path.normcase(os.path.abspath(p))
        if key in seen:
            continue
        seen.add(key)
        uniq.append(p)
    uniq.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return uniq


def _load_detailed_level_table_revenue(script_dir, data_android, data_ios, game_name=None):
    """Google / iOS ayri detailed_level_table veya purchase breakdown -> And/IOS Total $."""
    paths = _collect_revenue_level_table_paths(script_dir)
    if not paths:
        return pd.DataFrame(), pd.DataFrame(), ""

    print("   Purchase dosyalari:")
    for p in paths:
        plat = _detect_revenue_platform_from_filename(p)
        print(
            f"     - {_safe_console_text(os.path.basename(p))}  ->  "
            f"{_purchase_platform_display_label(plat)}"
        )

    and_parts = []
    ios_parts = []
    unlabeled = []

    for p in paths:
        plat = _detect_revenue_platform_from_filename(p)
        part = _read_revenue_level_table(p, game_name=game_name)
        if part.empty:
            continue
        if plat == "and":
            and_parts.append(part)
        elif plat == "ios":
            ios_parts.append(part)
        else:
            unlabeled.append(p)

    if unlabeled and not and_parts and not ios_parts and len(unlabeled) == 1:
        base = os.path.basename(unlabeled[0]).lower()
        if "detailed_level_table" in base:
            _print_purchase_export_naming_requirement(unlabeled)
        merged = _read_revenue_level_table(unlabeled[0], game_name=game_name)
        has_and = _platform_has_funnel_data(data_android)
        has_ios = _platform_has_funnel_data(data_ios)
        src = os.path.basename(unlabeled[0]) + " (platform etiketi yok; funnel'a gore)"
        return (
            merged.copy() if has_and else pd.DataFrame(),
            merged.copy() if has_ios else pd.DataFrame(),
            src,
        )

    if unlabeled and not and_parts and not ios_parts and len(unlabeled) != 2:
        _print_purchase_export_naming_requirement(unlabeled)
        print(
            f"   HATA: {len(unlabeled)} etiketsiz purchase dosyasi var; otomatik platform ayrimi yapilamadi.\n"
            "         Dosyalari google_/ios_ ile yeniden adlandirin veya analytics export adini duzeltsin.\n"
        )
        return pd.DataFrame(), pd.DataFrame(), ""

    if len(unlabeled) == 2 and not and_parts and not ios_parts:
        assigned = _prompt_revenue_file_platforms(unlabeled)
        if assigned.get("and"):
            and_parts.append(
                _read_revenue_level_table(assigned["and"], game_name=game_name)
            )
        if assigned.get("ios"):
            ios_parts.append(
                _read_revenue_level_table(assigned["ios"], game_name=game_name)
            )
        src = (
            f"And={os.path.basename(assigned['and'])}; "
            f"IOS={os.path.basename(assigned['ios'])}"
        )
        return _merge_revenue_parts(and_parts), _merge_revenue_parts(ios_parts), src

    if unlabeled:
        _warn_unlabeled_purchase_files(unlabeled, and_parts, ios_parts)

    and_df = _merge_revenue_parts(and_parts)
    ios_df = _merge_revenue_parts(ios_parts)

    if and_df.empty and ios_df.empty:
        return pd.DataFrame(), pd.DataFrame(), ""

    names = []
    if and_parts:
        names.append("And")
    if ios_parts:
        names.append("IOS")
    first = os.path.basename(paths[0]).lower()
    kind = (
        "detailed_level_table"
        if "detailed_level_table" in first
        else "purchase_breakdown"
    )
    src = kind + " (" + ", ".join(names) + ")"
    return and_df, ios_df, src


def _read_revenue_level_table_xlsx(path):
    """Geriye uyumluluk."""
    return _read_revenue_level_table(path)


def _load_revenue_from_csv(script_dir):
    """Eski akis: veri-*.csv (Store Name ile And/IOS ayrimi)."""
    csv_files = glob.glob(os.path.join(script_dir, "veri-*.csv"))
    csv_list = []
    for cf in csv_files:
        try:
            cf_df = pd.read_csv(cf)
            if all(k in cf_df.columns for k in ["Level", "USD Price", "Store Name"]):
                csv_list.append(cf_df)
        except Exception:
            pass
    if not csv_list:
        return pd.DataFrame(), pd.DataFrame(), ""
    all_csv = pd.concat(csv_list, ignore_index=True)
    all_csv["USD Price"] = pd.to_numeric(all_csv["USD Price"], errors="coerce").fillna(0)
    all_csv["Level"] = pd.to_numeric(all_csv["Level"], errors="coerce").dropna().astype(int)
    gu = all_csv[all_csv["Store Name"].astype(str).str.contains("Google", case=False, na=False)]
    csv_and = gu.groupby("Level")["USD Price"].sum().reset_index().rename(columns={"USD Price": "Total $"})
    ap = all_csv[all_csv["Store Name"].astype(str).str.contains("AppStore", case=False, na=False)]
    csv_ios = ap.groupby("Level")["USD Price"].sum().reset_index().rename(columns={"USD Price": "Total $"})
    names = ", ".join(os.path.basename(p) for p in csv_files[:3])
    if len(csv_files) > 3:
        names += f" (+{len(csv_files) - 3})"
    return csv_and, csv_ios, f"veri-*.csv ({names})"


def load_purchase_revenue(script_dir, data_android, data_ios, game_name=None):
    """
    And/IOS N (Total $) sutunu icin level bazli gelir.
    Oncelik: *detailed_level_table*.csv / *.xlsx
    Sonra: top_levels_by_purchase_count / purchase_count_game_breakdown CSV
    Yedek: veri-*.csv (Store Name ile And/IOS ayrimi).
    """
    and_df, ios_df, src = _load_detailed_level_table_revenue(
        script_dir, data_android, data_ios, game_name=game_name
    )
    if src:
        return and_df, ios_df, src
    return _load_revenue_from_csv(script_dir)


def print_quality_report_console(report):
    """Konsola kisa ozet."""
    print("\n--- Ozet ---")
    print(f"  Oyun: {report['game_name']} | Tarih: {report['date_range']}")
    for pk in ("Android", "iOS"):
        p = report["platforms"].get(pk, {})
        rf = p.get("recommended_funnel_level", p.get("recommended_level", 1))
        ra = p.get("recommended_ads_level", 1)
        hot = sorted(p.get("churn", {}).get("hotspots", []), key=lambda x: x["level"])
        top = ", ".join(f"L{x['level']}" for x in hot[:8]) if hot else "yok"
        print(f"  {pk}: funnel L{rf}+ | ADs L{ra}+ | churn uyarisi: {top}")
    insights = report.get("insights") or []
    if insights:
        print("  Insights:")
        for line in insights[:6]:
            print(f"   - {line}")
    print("  Detay tablolar: Excel > 'Ozet' sekmesi.\n")


def main():
    _configure_stdout_utf8()
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
    print(" Fur Agent — Otomatik Analiz baslatiliyor...")
    print("=======================================================\n")

    # 1. Veri dosyalarini SCRIPT_DIR icinde tara (oyun adi metaveriden okunur)
    print("1. Veri dosyalari (xlsx) taraniyor...")
    print(f"   Klasor: {os.path.normpath(SCRIPT_DIR)}")
    files, date_dict, game_dict, skipped = collect_data_xlsx_files(SCRIPT_DIR)

    if not files:
        print("HATA: Gecerli veri dosyasi bulunamadi (oyun metaverisi veya OS+Level+funnel yapisinda xlsx).")
        if skipped:
            print("   Atlanan dosyalar (ilk 20):")
            for name, why in skipped[:20]:
                print(f"     - {name}: {why}")
            if len(skipped) > 20:
                print(f"     ... ve {len(skipped) - 20} dosya daha")
        sys.exit(1)

    detected_game = _resolve_detected_game_name(game_dict)
    if detected_game:
        print(f"   OK Oyun: {detected_game}")
    else:
        detected_game = "Hole_Pool"
        print(
            "   UYARI: Oyun metaverisi bulunamadi (# Farm Block Escape gibi); "
            "cikti adi varsayilan Hole_Pool olacak."
        )

    max_level = get_max_level_from_user()
    print(f"\n[OK] Analiz {max_level}. levele kadar yapilacak.\n")

    if skipped:
        print(f"   Bilgi: {len(skipped)} dosya atlandi (sablon/cikti veya veri degil).")
    print(f"   OK {len(files)} adet veri dosyasi tespit edildi.")
    unique_dates = set(date_dict.values())
    if len(unique_dates) > 1:
        print(f"   UYARI: Tarih farkliligi var: {unique_dates}")
    elif unique_dates:
        print(f"   OK Tarih: {list(unique_dates)[0]}\n")

    # 2. Verileri oku ve funnel tipine gore ayristir
    print("2. Veriler okunuyor...")
    data_android = {}  # { funnel_type: list[df] }
    data_ios = {}
    data_unspecified = {}
    missing_platform_files = []

    for f in files:
        funnel, df_and, df_ios, has_platform = read_platform_data(f)
        n_and = 0 if df_and is None or df_and.empty else len(df_and)
        n_ios = 0 if df_ios is None or df_ios.empty else len(df_ios)
        if has_platform:
            print(f"   {os.path.basename(f)} -> {funnel} (And {n_and} / IOS {n_ios})")
        else:
            n_all = n_and
            print(
                f"   {os.path.basename(f)} -> {funnel} "
                f"({n_all} satir, Platform yok)"
            )
        if has_platform:
            if funnel not in data_android:
                data_android[funnel] = []
            if funnel not in data_ios:
                data_ios[funnel] = []
            data_android[funnel].append(df_and)
            data_ios[funnel].append(df_ios)
        else:
            if funnel not in data_unspecified:
                data_unspecified[funnel] = []
            data_unspecified[funnel].append(df_and)
            missing_platform_files.append(os.path.basename(f))

    if data_unspecified:
        target = get_unspecified_platform_target_from_user(missing_platform_files)
        label = {"and": "And", "ios": "IOS", "both": "And+IOS"}[target]
        print(f"   -> Platform'siz veri '{label}' sekmesine yazilacak.\n")
        for ftype, dfs in data_unspecified.items():
            if target in ("and", "both"):
                data_android.setdefault(ftype, []).extend(dfs)
            if target in ("ios", "both"):
                copies = [d.copy() for d in dfs] if target == "both" else dfs
                data_ios.setdefault(ftype, []).extend(copies)

    # Aynı funnel tipindeki dosyalari birlestir (Eger varsa)
    for ftype in list(data_android.keys()):
        data_android[ftype] = safe_concat(data_android[ftype])
    for ftype in list(data_ios.keys()):
        data_ios[ftype] = safe_concat(data_ios[ftype])

    # 3. Gelir verileri (oncelik: detailed_level_table xlsx, yedek: veri-*.csv)
    print("3. Gelir (revenue) verileri okunuyor...")
    csv_and_usd, csv_ios_usd, revenue_src = load_purchase_revenue(
        SCRIPT_DIR, data_android, data_ios, detected_game
    )
    if revenue_src:
        print(f"   OK Gelir kaynagi: {revenue_src}")
    else:
        print(
            "   Bilgi: Gelir dosyasi bulunamadi "
            "(detailed_level_table, top_levels_by_purchase veya veri-*.csv)."
        )

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

    metric_labels = _extract_and_ios_metric_labels(wb)
    quality_report = build_data_quality_report(
        data_android,
        data_ios,
        game_name=detected_game,
        date_range=date_range,
        country_code=country_code,
        max_level=max_level,
        source_files=files,
        revenue_android_df=csv_and_usd,
        revenue_ios_df=csv_ios_usd,
        revenue_source=revenue_src,
        metric_labels=metric_labels,
    )
    workbook_ctx = _extract_workbook_insight_context(wb, max_level)
    quality_report = _apply_workbook_context_to_report(quality_report, workbook_ctx)
    quality_report["insights"] = _build_auto_insights_from_workbook(quality_report)
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
