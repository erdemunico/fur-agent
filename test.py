import pandas as pd
import glob


def dedupe_columns_keep_left(df):
    """Aynı isimli sütunlarda ilk (soldaki) sütunu bırakır; concat için gerekli."""
    seen = set()
    keep_idx = []
    for i, col in enumerate(df.columns):
        if col in seen:
            continue
        seen.add(col)
        keep_idx.append(i)
    return df.iloc[:, keep_idx].copy()


def drop_blank_named_columns(df):
    """Başlığı boş veya eksik (NaN) olan sütunları kaldırır; sadece bu scriptteki birleşimi sadeleştirir."""

    def _is_blank(name):
        if isinstance(name, str):
            return name.strip() == ""
        try:
            return bool(pd.isna(name))
        except (TypeError, ValueError):
            return False

    to_drop = [c for c in df.columns if _is_blank(c)]
    if not to_drop:
        return df
    return df.drop(columns=to_drop)


all_android = []
all_ios = []

files = glob.glob('Free form*.xlsx')
for f in files:
    print(f"\nProcessing {f}")
    df_raw = pd.read_excel(f, header=None)
    
    os_row_idx = None
    level_row_idx = None
    
    for idx, row in df_raw.head(15).iterrows():
        row_list = list(row.dropna().astype(str))
        if any('Operating system' in str(c) for c in row):
            os_row_idx = idx
        if any('Level' in str(c) for c in row):
            level_row_idx = idx

    print(f"OS row: {os_row_idx}, Level row: {level_row_idx}")
    
    # Check if OS is a pivot row or a column header
    os_row_vals = df_raw.iloc[os_row_idx].fillna('').astype(str).str.strip().tolist()
    
    if 'Android' in os_row_vals or 'iOS' in os_row_vals:
        print(" -> Detected Pivot Layout")
        level_col_idx = df_raw.iloc[level_row_idx].tolist().index('Level')
        
        # Build Android df
        and_cols = [level_col_idx] + [i for i, v in enumerate(os_row_vals) if v == 'Android']
        df_and = df_raw.iloc[level_row_idx + 1:, and_cols].copy()
        df_and.columns = df_raw.iloc[level_row_idx, and_cols].tolist()
        df_and = dedupe_columns_keep_left(df_and)
        df_and = drop_blank_named_columns(df_and)
        df_and['SourceFile'] = f
        all_android.append(df_and)
        
        # Build iOS df
        ios_cols = [level_col_idx] + [i for i, v in enumerate(os_row_vals) if v == 'iOS']
        df_ios = df_raw.iloc[level_row_idx + 1:, ios_cols].copy()
        df_ios.columns = df_raw.iloc[level_row_idx, ios_cols].tolist()
        df_ios = dedupe_columns_keep_left(df_ios)
        df_ios = drop_blank_named_columns(df_ios)
        df_ios['SourceFile'] = f
        all_ios.append(df_ios)
        
    else:
        print(" -> Detected Flat Layout")
        # Header is level_row_idx
        df_flat = df_raw.iloc[level_row_idx + 1:].copy()
        df_flat.columns = df_raw.iloc[level_row_idx].tolist()
        df_flat = dedupe_columns_keep_left(df_flat)
        df_flat = df_flat.dropna(subset=['Level']) # filter out grand total if any?
        # Actually grand total might not have level.
        df_flat['SourceFile'] = f
        
        # Filter
        df_and = df_flat[df_flat['Operating system'].astype(str).str.contains('Android', case=False, na=False)].copy()
        df_ios = df_flat[df_flat['Operating system'].astype(str).str.contains('iOS', case=False, na=False)].copy()
        
        df_and = df_and.drop(columns=['Operating system'], errors='ignore')
        df_ios = df_ios.drop(columns=['Operating system'], errors='ignore')
        df_and = drop_blank_named_columns(df_and)
        df_ios = drop_blank_named_columns(df_ios)

        all_android.append(df_and)
        all_ios.append(df_ios)

df_android = pd.concat(all_android, ignore_index=True)
df_ios = pd.concat(all_ios, ignore_index=True)

print(f"Android Total Rows: {len(df_android)}, Columns: {df_android.columns.tolist()}")
print(f"iOS Total Rows: {len(df_ios)}, Columns: {df_ios.columns.tolist()}")

