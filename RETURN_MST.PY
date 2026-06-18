import pandas as pd
import win32com.client as win32
import os

# =========================
# PATHS
# =========================
source_file = r"C:\Users\Dilip\OneDrive - Happy Ecom Venture\Desktop\MASTER RETURN FILE.xlsx"
output_file = r"C:\Users\Dilip\Downloads\Return master lsd.xlsx"
temp_folder = r"C:\Users\Dilip\Downloads\temp_mail"

os.makedirs(temp_folder, exist_ok=True)

# =========================
# LOAD MASTER
# =========================
df = pd.read_excel(source_file)
df.columns = df.columns.str.strip()

required_cols = [
    "Tracking ID", "Date", "Dilip Remarks",
    "NSL First", "NSL Last",
    "MP First", "MP Last",
    "MD First", "MD Last",
    "Wrong Scan Date",
    "Current Status"
]

for col in required_cols:
    if col not in df.columns:
        df[col] = (
            pd.NaT
            if ("Date" in col or "First" in col or "Last" in col)
            else ""
        )

df["Tracking ID"] = df["Tracking ID"].astype(str).str.strip()
df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

# =========================
# OUTLOOK READ
# =========================
outlook = win32.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
sent = namespace.GetDefaultFolder(5)

messages = sent.Items
messages.Sort("[SentOn]", True)

mail_data = {}

for msg in messages:

    try:

        # Only Mail Items
        if msg.Class != 43:
            continue

        # Subject Check (Flexible)
        subject = str(msg.Subject).upper().strip()

        if "RETURNS" not in subject or "MAPPING" not in subject:
            continue

        if msg.SentOn is None:
            continue

        mail_date = pd.to_datetime(msg.SentOn).tz_localize(None)

        print(f"\n📩 READING MAIL : {subject}")
        print(f"📅 MAIL DATE    : {mail_date}")

        # =========================
        # ATTACHMENTS
        # =========================
        for att in msg.Attachments:

            file_name = att.FileName.lower()

            # Read XLS + XLSX
            if not (
                file_name.endswith(".xlsx")
                or file_name.endswith(".xls")
            ):
                continue

            print(f"📎 ATTACHMENT   : {att.FileName}")

            # Proper temp save
            temp_file = os.path.join(
                temp_folder,
                f"{msg.EntryID}_{att.FileName}"
            )

            try:
                att.SaveAsFile(temp_file)
                xls = pd.ExcelFile(temp_file)

            except Exception as e:
                print("❌ FILE OPEN ERROR:", e)
                continue

            # =========================
            # SHEET MAP
            # =========================
            sheet_map = {
                "NEED SHIPPING LABEL": "NSL",
                "MAPPING PENDING": "MP",
                "MAPPED DONE": "MD",
                "WRONG SCAN": "WS"
            }

            # =========================
            # SHEET LOOP
            # =========================
            for sheet in xls.sheet_names:

                for key in sheet_map:

                    if key in sheet.upper():

                        print(f"✅ SHEET READ    : {sheet}")

                        try:
                            temp_df = pd.read_excel(
                                xls,
                                sheet_name=sheet
                            )

                        except Exception as e:
                            print("❌ SHEET ERROR:", e)
                            continue

                        # Find tracking column
                        tcol = next(
                            (
                                c for c in temp_df.columns
                                if "tracking" in str(c).lower()
                            ),
                            None
                        )

                        if not tcol:
                            print("❌ TRACKING COLUMN NOT FOUND")
                            continue

                        temp_df.rename(
                            columns={tcol: "Tracking ID"},
                            inplace=True
                        )

                        temp_df["Tracking ID"] = (
                            temp_df["Tracking ID"]
                            .astype(str)
                            .str.strip()
                        )

                        # =========================
                        # STORE MAIL DATA
                        # =========================
                        for tid in temp_df["Tracking ID"]:

                            if tid not in mail_data:

                                mail_data[tid] = {
                                    "NSL": [],
                                    "MP": [],
                                    "MD": [],
                                    "WS": []
                                }

                            mail_data[tid][sheet_map[key]].append(mail_date)

            # Delete temp file
            try:
                os.remove(temp_file)

            except:
                pass

    except Exception as e:
        print("❌ MAIL ERROR:", e)

# =========================
# GLOBAL LATEST
# =========================
# WS excluded intentionally
all_dates = []

for tid in mail_data:
    for k in ["NSL", "MP", "MD"]:
        all_dates += mail_data[tid][k]

GLOBAL_LATEST = max(all_dates) if all_dates else None
GLOBAL_LATEST_DATE = (
    GLOBAL_LATEST.date()
    if GLOBAL_LATEST
    else None
)

print("\n=========================")
print("🌍 GLOBAL LATEST DATE :", GLOBAL_LATEST_DATE)
print("=========================\n")

# =========================
# FINAL LOGIC
# =========================
cutoff = pd.to_datetime("2026-03-01")

for i, row in df.iterrows():

    tid = row["Tracking ID"]
    remarks = str(row["Dilip Remarks"]).upper()

    # =========================
    # MASTER WRONG SCAN
    # =========================
    if "WRONG SCAN" in remarks:

        df.loc[i, "Wrong Scan Date"] = row["Date"]
        df.loc[i, "Current Status"] = "REMOVE DONE"
        continue

    # =========================
    # OLD DATA
    # =========================
    if pd.notna(row["Date"]) and row["Date"] < cutoff:

        df.loc[i, "Current Status"] = "DATA NOT RECOVER"
        continue

    # =========================
    # NOT FOUND
    # =========================
    if tid not in mail_data or GLOBAL_LATEST is None:

        df.loc[i, "Current Status"] = "QC DONE / CLOSED"
        continue

    data = mail_data[tid]

    # =========================
    # UPDATE DATE COLUMNS
    # =========================
    for r in ["NSL", "MP", "MD"]:

        if data[r]:

            df.loc[i, f"{r} First"] = min(data[r])
            df.loc[i, f"{r} Last"] = max(data[r])

    latest_nsl = max(data["NSL"]) if data["NSL"] else None
    latest_mp  = max(data["MP"]) if data["MP"] else None
    latest_md  = max(data["MD"]) if data["MD"] else None
    latest_ws  = max(data["WS"]) if data["WS"] else None

    # =========================
    # WRONG SCAN LOGIC
    # =========================
    if latest_ws:

        tracking_latest = max(
            [
                d for d in
                [latest_nsl, latest_mp, latest_md, latest_ws]
                if d is not None
            ],
            default=None
        )

        if latest_ws == tracking_latest:

            df.loc[i, "Wrong Scan Date"] = latest_ws
            df.loc[i, "Current Status"] = "REMOVE DONE"
            continue

    # =========================
    # TRACKING LATEST
    # =========================
    tracking_latest = max(
        [
            d for d in
            [latest_nsl, latest_mp, latest_md]
            if d is not None
        ],
        default=None
    )

    # =========================
    # NOT IN GLOBAL LATEST
    # =========================
    in_latest = (
        tracking_latest
        and tracking_latest.date() == GLOBAL_LATEST_DATE
    )

    if not in_latest:

        df.loc[i, "Current Status"] = "QC DONE / CLOSED"
        continue

    # =========================
    # NEED SHIPPING LABEL
    # =========================
    if latest_nsl and latest_nsl.date() == GLOBAL_LATEST_DATE:

        if min(data["NSL"]) == max(data["NSL"]):

            df.loc[i, "Current Status"] = "IN PROCESS"

        else:

            df.loc[i, "Current Status"] = "PENDING FROM WH"

        continue

    # =========================
    # MAPPING PENDING
    # =========================
    if latest_mp and latest_mp.date() == GLOBAL_LATEST_DATE:

        if min(data["MP"]) == max(data["MP"]):

            df.loc[i, "Current Status"] = "IN PROCESS"

        elif latest_mp > (latest_nsl or pd.Timestamp.min):

            df.loc[i, "Current Status"] = "PENDING FROM HEC"

        else:

            df.loc[i, "Current Status"] = "MAPPING PENDING"

        continue

    # =========================
    # MAPPED DONE
    # =========================
    if latest_md and latest_md.date() == GLOBAL_LATEST_DATE:

        df.loc[i, "Current Status"] = "QC DONE / CLOSED"
        continue

    # =========================
    # DEFAULT
    # =========================
    df.loc[i, "Current Status"] = "QC DONE / CLOSED"

# =========================
# SAVE
# =========================
df.to_excel(output_file, index=False)

print("\n✅ FINAL CODE EXECUTED SUCCESSFULLY")
print(f"📁 OUTPUT FILE : {output_file}")