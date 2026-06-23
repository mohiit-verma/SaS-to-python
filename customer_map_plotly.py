# ── CELL 1: Imports & file paths ─────────────────────────────────────────────

import pandas as pd

DB_FILE  = "databricks_export.csv"   # change to your file name (.csv or .xlsx)
SAS_FILE = "sas_output.csv"          # change to your file name (.csv or .xlsx)


# ── CELL 2: Load files ────────────────────────────────────────────────────────

def load(path):
    if path.endswith(".csv"):
        df = pd.read_csv(path)
    else:
        df = pd.read_excel(path)
    df.columns = df.columns.str.strip().str.lower()
    return df

df_db  = load(DB_FILE)
df_sas = load(SAS_FILE)

print("Databricks :", df_db.shape)
print("SAS        :", df_sas.shape)


# ── CELL 3: Sort both on all columns ─────────────────────────────────────────

common_cols = sorted(set(df_db.columns) & set(df_sas.columns))

df_db  = df_db[common_cols].sort_values(by=common_cols).reset_index(drop=True)
df_sas = df_sas[common_cols].sort_values(by=common_cols).reset_index(drop=True)

print("Sorted on", len(common_cols), "columns")


# ── CELL 4: Compare row by row ────────────────────────────────────────────────

# Returns a True/False table — True means the cell is different
diff_mask = df_db.astype(str) != df_sas.astype(str)

print("Total rows compared  :", len(df_db))
print("Rows with any diff   :", diff_mask.any(axis=1).sum())
print("Columns with any diff:", diff_mask.any(axis=0).sum())


# ── CELL 5: See which rows differ ────────────────────────────────────────────

differing_rows = diff_mask[diff_mask.any(axis=1)]
print(differing_rows)


# ── CELL 6: See side-by-side for a specific column ───────────────────────────

# Change "column_name" to whichever column you want to inspect
col = "column_name"

diff_in_col = diff_mask[col]
print(pd.DataFrame({
    "row"      : diff_in_col[diff_in_col].index,
    "db_value" : df_db.loc[diff_in_col, col].values,
    "sas_value": df_sas.loc[diff_in_col, col].values,
}))


# ── CELL 7: Export all differences ───────────────────────────────────────────

diff_mask.to_csv("diff_mask.csv")
print("Saved diff_mask.csv")