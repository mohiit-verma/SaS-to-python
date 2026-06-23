# ============================================================
#  Dataset Comparison: Databricks vs SAS
#  Run cell-by-cell in a Jupyter / Python notebook
# ============================================================

# ── CELL 1: Imports & config ─────────────────────────────────────────────────

import pandas as pd
import numpy as np
from pathlib import Path

# ┌─────────────────────────────────────────────────────────┐
# │  UPDATE THESE PATHS AND SETTINGS BEFORE RUNNING         │
# └─────────────────────────────────────────────────────────┘

DB_FILE   = "databricks_export.csv"   # or .xlsx
SAS_FILE  = "sas_output.csv"          # or .xlsx

KEY_COLS  = ["account_id"]            # primary key column(s) — change as needed
NUM_TOL   = 1e-6                      # tolerance for numeric comparisons

# Output folder for mismatch CSVs (created automatically)
OUTPUT_DIR = Path("comparison_output")
OUTPUT_DIR.mkdir(exist_ok=True)


# ── CELL 2: Load data ─────────────────────────────────────────────────────────

def load_file(path: str) -> pd.DataFrame:
    """Load CSV or Excel file into a DataFrame."""
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"File not found: {path}")
    if p.suffix.lower() in [".xlsx", ".xls"]:
        df = pd.read_excel(p)
    elif p.suffix.lower() == ".csv":
        df = pd.read_csv(p)
    else:
        raise ValueError(f"Unsupported file type: {p.suffix}")
    # Normalise column names — lowercase, strip spaces
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
    return df

df_db  = load_file(DB_FILE)
df_sas = load_file(SAS_FILE)

print(f"Databricks : {df_db.shape[0]:,} rows × {df_db.shape[1]} columns")
print(f"SAS        : {df_sas.shape[0]:,} rows × {df_sas.shape[1]} columns")


# ── CELL 3: Schema comparison ─────────────────────────────────────────────────

cols_db  = set(df_db.columns)
cols_sas = set(df_sas.columns)
common   = cols_db & cols_sas

print("=" * 55)
print("SCHEMA CHECK")
print("=" * 55)
print(f"  Total columns — DB: {len(cols_db)}, SAS: {len(cols_sas)}, Common: {len(common)}")

only_db  = cols_db  - cols_sas
only_sas = cols_sas - cols_db
if only_db:
    print(f"\n  Only in Databricks ({len(only_db)}): {sorted(only_db)}")
if only_sas:
    print(f"\n  Only in SAS ({len(only_sas)}): {sorted(only_sas)}")
if not only_db and not only_sas:
    print("\n  Column names match exactly ✓")

# Data-type comparison for common columns
dtype_df = pd.DataFrame({
    "databricks_dtype": df_db[sorted(common)].dtypes,
    "sas_dtype":        df_sas[sorted(common)].dtypes,
})
dtype_df["match"] = dtype_df["databricks_dtype"].astype(str) == dtype_df["sas_dtype"].astype(str)
type_mismatches = dtype_df[~dtype_df["match"]]
if not type_mismatches.empty:
    print("\n  Data-type mismatches (usually safe to ignore):")
    print(type_mismatches.to_string())
else:
    print("\n  Data types match for all common columns ✓")


# ── CELL 4: Profile (nulls, duplicates, stats) ────────────────────────────────

def profile(df: pd.DataFrame, label: str):
    print(f"\n{'─'*55}")
    print(f"  {label}")
    print(f"{'─'*55}")
    print(f"  Rows        : {len(df):,}")
    print(f"  Columns     : {len(df.columns)}")
    print(f"  Duplicates  : {df.duplicated(subset=[k for k in KEY_COLS if k in df.columns]).sum():,}")

    null_counts = df.isnull().sum()
    null_cols   = null_counts[null_counts > 0]
    if null_cols.empty:
        print("  Nulls       : none")
    else:
        print(f"  Nulls       :\n{null_cols.to_string()}")

    numeric_cols = df.select_dtypes(include="number").columns
    if len(numeric_cols):
        print(f"\n  Numeric summary ({len(numeric_cols)} columns):")
        print(df[numeric_cols].describe().round(4).to_string())

print("=" * 55)
print("DATA PROFILE")
print("=" * 55)
profile(df_db,  "Databricks")
profile(df_sas, "SAS")


# ── CELL 5: Row-count & duplicate check ───────────────────────────────────────

print("=" * 55)
print("ROW COUNT & DUPLICATE SUMMARY")
print("=" * 55)

rows_match = len(df_db) == len(df_sas)
print(f"  Row count match : {'YES ✓' if rows_match else 'NO ✗'} "
      f"(DB={len(df_db):,}, SAS={len(df_sas):,})")

key_present = all(k in df_db.columns and k in df_sas.columns for k in KEY_COLS)
if key_present:
    dup_db  = df_db.duplicated(subset=KEY_COLS).sum()
    dup_sas = df_sas.duplicated(subset=KEY_COLS).sum()
    print(f"  Duplicates on key — DB: {dup_db:,}, SAS: {dup_sas:,}")
else:
    print(f"  Key column(s) {KEY_COLS} not found in both datasets — skipping duplicate check.")


# ── CELL 6: Merge & row matching ─────────────────────────────────────────────

key_present = all(k in df_db.columns and k in df_sas.columns for k in KEY_COLS)
if not key_present:
    raise ValueError(f"Key column(s) {KEY_COLS} missing. Update KEY_COLS in Cell 1.")

merged = df_db.merge(
    df_sas,
    on=KEY_COLS,
    how="outer",
    suffixes=("_db", "_sas"),
    indicator=True,
)

merge_counts = merged["_merge"].value_counts()
both         = int(merge_counts.get("both",       0))
only_left    = int(merge_counts.get("left_only",  0))
only_right   = int(merge_counts.get("right_only", 0))

print("=" * 55)
print("ROW MATCHING")
print("=" * 55)
print(f"  Matched (both)          : {both:,}")
print(f"  Only in Databricks      : {only_left:,}")
print(f"  Only in SAS             : {only_right:,}")

# Export unmatched rows
if only_left > 0:
    path = OUTPUT_DIR / "rows_only_in_databricks.csv"
    merged[merged["_merge"] == "left_only"].drop(columns="_merge").to_csv(path, index=False)
    print(f"\n  Saved → {path}")

if only_right > 0:
    path = OUTPUT_DIR / "rows_only_in_sas.csv"
    merged[merged["_merge"] == "right_only"].drop(columns="_merge").to_csv(path, index=False)
    print(f"  Saved → {path}")


# ── CELL 7: Column-by-column value comparison ─────────────────────────────────

matched = merged[merged["_merge"] == "both"].copy()

# Columns to compare (exclude key cols)
compare_cols = [
    c for c in common
    if c not in KEY_COLS
    and f"{c}_db"  in matched.columns
    and f"{c}_sas" in matched.columns
]

mismatch_summary = {}

for col in compare_cols:
    a = matched[f"{col}_db"]
    b = matched[f"{col}_sas"]

    if pd.api.types.is_numeric_dtype(a) and pd.api.types.is_numeric_dtype(b):
        diff_mask = (a - b).abs() > NUM_TOL
    else:
        diff_mask = a.astype(str).str.strip() != b.astype(str).str.strip()

    n_diff = int(diff_mask.sum())
    if n_diff > 0:
        mismatch_summary[col] = n_diff

print("=" * 55)
print("VALUE-LEVEL COMPARISON (matched rows only)")
print("=" * 55)

if not mismatch_summary:
    print("  All values match across all common columns ✓")
else:
    print(f"  {len(mismatch_summary)} column(s) have mismatches:\n")
    for col, cnt in sorted(mismatch_summary.items(), key=lambda x: -x[1]):
        pct = cnt / len(matched) * 100
        print(f"  {col:<35} {cnt:>6,} rows  ({pct:.1f}%)")

    # Export mismatch detail for each column
    print("\n  Exporting detail files...")
    for col in mismatch_summary:
        a = matched[f"{col}_db"]
        b = matched[f"{col}_sas"]
        if pd.api.types.is_numeric_dtype(a) and pd.api.types.is_numeric_dtype(b):
            diff_mask = (a - b).abs() > NUM_TOL
        else:
            diff_mask = a.astype(str).str.strip() != b.astype(str).str.strip()

        detail = matched.loc[diff_mask, KEY_COLS + [f"{col}_db", f"{col}_sas"]].copy()
        if pd.api.types.is_numeric_dtype(a):
            detail[f"{col}_diff"] = (a[diff_mask] - b[diff_mask]).round(6)

        path = OUTPUT_DIR / f"mismatch_{col}.csv"
        detail.to_csv(path, index=False)
        print(f"    Saved → {path}")


# ── CELL 8: Summary report ───────────────────────────────────────────────────

print()
print("=" * 55)
print("FINAL SUMMARY REPORT")
print("=" * 55)
print(f"  Databricks file  : {DB_FILE}")
print(f"  SAS file         : {SAS_FILE}")
print(f"  Key column(s)    : {KEY_COLS}")
print()
print(f"  Row count match  : {'YES ✓' if rows_match else 'NO ✗'}")
print(f"  Schema match     : {'YES ✓' if not only_db and not only_sas else 'NO ✗'}")
print(f"  Matched rows     : {both:,}")
print(f"  Unmatched (DB)   : {only_left:,}")
print(f"  Unmatched (SAS)  : {only_right:,}")
print(f"  Columns checked  : {len(compare_cols)}")
print(f"  Columns clean    : {len(compare_cols) - len(mismatch_summary)}")
print(f"  Columns mismatch : {len(mismatch_summary)}")
print()
if mismatch_summary:
    print("  Columns with mismatches:")
    for col, cnt in sorted(mismatch_summary.items(), key=lambda x: -x[1]):
        print(f"    • {col}: {cnt:,} rows differ")
print()
print(f"  Output files saved to: ./{OUTPUT_DIR}/")
print("=" * 55)