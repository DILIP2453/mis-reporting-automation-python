import pandas as pd
import re
from datetime import datetime

# CLEAN FUNCTION
def clean_text(x):
    if pd.isna(x):
        return ""
    x = str(x)
    x = x.replace('\xa0', '')
    x = re.sub(r'\s+', '', x)
    return x.strip().upper()

# FILE PATHS
input_file = r"C:\Users\Dilip\Downloads\Myntra 17-06-2026.xlsx"
mis_file = r"C:\Users\Dilip\OneDrive - Happy Ecom Venture\Desktop\ALL LOC MIS.xlsx"

now = datetime.now().strftime("%d-%m-%Y_%H-%M")
output_file = fr"C:\Users\Dilip\Downloads\Myntra_output_{now}.xlsx"

print("🚀 Loading MIS file...")

# LOAD MIS FILE
all_sheets = pd.read_excel(mis_file, sheet_name=None)

combined_data = []

for sheet_name, df in all_sheets.items():
    df.columns = [str(c).strip() for c in df.columns]

    if len(df.columns) <= 5:
        continue

    tracking_col = df.columns[5]

    temp = df[[tracking_col]].copy()
    temp.columns = ["Tracking_ID"]

    temp["Tracking_ID"] = temp["Tracking_ID"].map(clean_text)
    temp = temp[temp["Tracking_ID"] != ""]

    temp["Sheet_Name"] = sheet_name

    combined_data.append(temp)

combined_df = pd.concat(combined_data, ignore_index=True)

# LOOKUP DICT
lookup_dict = dict(zip(combined_df["Tracking_ID"], combined_df["Sheet_Name"]))

print("✅ MIS Ready")

# LOAD INPUT
input_df = pd.read_excel(input_file)
input_df.columns = [str(c).strip() for c in input_df.columns]

col_W = input_df.columns[22]
col_X = input_df.columns[23]
col_Y = input_df.columns[24]

# CLEAN INPUT
input_df[col_W] = input_df[col_W].map(clean_text)
input_df[col_X] = input_df[col_X].map(clean_text)
input_df[col_Y] = input_df[col_Y].map(clean_text)

print("⚡ Matching started...")

# =========================
# FAST MATCH (NO APPLY)
# =========================

# Map each column
match_W = input_df[col_W].map(lookup_dict)
match_X = input_df[col_X].map(lookup_dict)
match_Y = input_df[col_Y].map(lookup_dict)

# First match logic (same as IFERROR chain)
input_df["Found_In_Sheet"] = match_W.fillna(match_X).fillna(match_Y)

# Tracking ID return (same logic)
input_df["Matched_Tracking_ID"] = (
    input_df[col_W].where(match_W.notna())
    .fillna(input_df[col_X].where(match_X.notna()))
    .fillna(input_df[col_Y].where(match_Y.notna()))
)

# Fill Not Found
input_df["Found_In_Sheet"] = input_df["Found_In_Sheet"].fillna("Not Found")
input_df["Matched_Tracking_ID"] = input_df["Matched_Tracking_ID"].fillna("Not Found")

print("✅ Matching Done")

# SAVE
input_df.to_excel(output_file, index=False)

print("💾 File saved:", output_file)