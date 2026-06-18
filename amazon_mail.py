import pandas as pd
import win32com.client as win32
import os
import traceback

# =========================
# SETTINGS
# =========================
FILE_PATH = r"C:\Users\Dilip\Downloads\Amazon_output_12-06-2026_16-53.xlsx"
SHEET_NAME = "Succesfull delivered"
LOGO_PATH = r"C:\Users\Dilip\Downloads\Outlook-5uukthm0.png"

START_DATE = "2026-06-12"
END_DATE = "2026-06-14"   # 👈 IMPORTANT: adjust if file is June data

# =========================
# READ FILE
# =========================
df = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME)

df.columns = df.columns.astype(str).str.strip()

print("✅ File Loaded")
print("Total Rows:", len(df))

# =========================
# REQUIRED COLUMNS CHECK
# =========================
required_cols = [
    "Date",
    "Location",
    "Order Id",
    "Tracking Number",
    "Found_In_Sheet"
]

df = df[required_cols]

# =========================
# DATE CLEAN (VERY IMPORTANT FIX)
# =========================
df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
df = df.dropna(subset=["Date"])

# Time remove
df["Date"] = df["Date"].dt.date

start_date = pd.to_datetime(START_DATE).date()
end_date = pd.to_datetime(END_DATE).date()

df = df[
    (df["Date"] >= start_date) &
    (df["Date"] <= end_date)
]

print("\nDates After Filter:")
print(sorted(df["Date"].astype(str).unique()))
print("Rows After Date Filter:", len(df))

print("Rows After Date Filter:", len(df))

# =========================
# CLEAN FOUND COLUMN
# =========================
df["Found_In_Sheet"] = (
    df["Found_In_Sheet"]
    .astype(str)
    .str.strip()
    .str.upper()
)

print("\nUnique Found_In_Sheet Values:")
print(df["Found_In_Sheet"].unique())

# =========================
# FILTER NOT FOUND
# =========================
df = df[df["Found_In_Sheet"] == "NOT FOUND"]

print("Rows After NOT FOUND Filter:", len(df))

# =========================
# STOP IF EMPTY
# =========================
if df.empty:
    print("\n❌ No NOT FOUND Data Available in selected date range")
    input("Press Enter to exit...")
    raise SystemExit

# =========================
# EMAIL MAPPING
# =========================
warehouse_email_map = {
    "BWD": {
        "TO": ["sameer.bhoir@maersk.com", "nilesh.bajage@apmterminals.com"],
        "CC": ["chetan@happyecom.com"]
    },
    "KOL": {
        "TO": ["tutu.hazra@apmterminals.com"],
        "CC": ["tuhin@happyecom.com"]
    },
    "GGN": {
        "TO": ["ramesh.kumar@apmterminals.com", "nitin.sharma@apmterminals.com"],
        "CC": ["alok@happyecom.com", "naven.singh@maersk.com"]
    },
    "BNG": {
        "TO": ["padmavathamma.sr3@flipkart.com", "basvaraj@happyecom.com"],
        "CC": ["sivakumar.v1@flipkart.com"]
    },
    "AHMD": {
        "TO": ["ops.ahmedabad@emizainc.com"],
        "CC": []
    },
    "GUW": {
        "TO": [
            "debarup.chanda@apmterminals.com",
            "chetan@happyecom.com"
        ],
        "CC": [
            "Sumit.boro@apmterminals.com",
            "ajeet.kumar.sharma@maersk.com",
            "ravi.ranjan.4@maersk.com"
        ]
   },

        "HYD": {
    "TO": [
        "v.teja@emizainc.com"
    ],
    "CC": []
    }
}

COMMON_CC = [
    "kunal@happyecom.com",
    "samir@happyecom.com"
]

# =========================
# OUTLOOK
# =========================
outlook = win32.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
sent_folder = namespace.GetDefaultFolder(5)

# =========================
# SIGNATURE
# =========================
def build_signature():
    return """
    <br><br>
    Regards,<br>
    <b>Dilip Patekar</b><br>
    MIS Associate<br><br>
    <img src="cid:logoimg" width="120">
    """

# =========================
# GROUP BY LOCATION
# =========================
grouped = df.groupby("Location")

# =========================
# PROCESS
# =========================
for wh, data in grouped:

    try:
        wh = str(wh).strip().upper()

        print("\n======================")
        print("Location:", wh)
        print("Cases:", len(data))

        if wh not in warehouse_email_map:
            print("❌ No email mapping")
            continue

        # =========================
        # CREATE EXCEL
        # =========================
        file_name = f"AMAZON_{wh}_NOT_FOUND.xlsx"
        file_path = os.path.join(os.getcwd(), file_name)

        data.to_excel(file_path, index=False)

        # =========================
        # SUBJECT
        # =========================
        subject_line = f"AMAZON | {wh} | Shipment Not Received at Warehouse"

        # =========================
        # OLD MAIL SEARCH
        # =========================
        mail = None
        messages = sent_folder.Items
        messages.Sort("[SentOn]", True)

        for msg in messages:
            try:
                if msg.Class != 43:
                    continue

                subj = str(msg.Subject).lower()

                if wh.lower() in subj and "amazon" in subj:
                    mail = msg.ReplyAll()
                    break
            except:
                pass

        # =========================
        # NEW MAIL
        # =========================
        if mail is None:
            mail = outlook.CreateItem(0)
            mail.Subject = subject_line

        # =========================
        # BODY
        # =========================
        body = f"""
Hi Team,<br><br>

We have observed that the below shipments are marked as delivered on the marketplace; however, the same has not been received at the warehouse.<br><br>

Request you to kindly check and confirm the status from your end.<br><br>

<b>Total Cases:</b> {len(data)}<br>
"""

        # =========================
        # EMAILS
        # =========================
        to_list = warehouse_email_map[wh]["TO"].copy()
        cc_list = warehouse_email_map[wh]["CC"].copy() + COMMON_CC

        if wh == "BNG":
            to_list.append("bala.murali@flipkart.com")
            cc_list.append("syed.yousuf@flipkart.com")

        mail.To = ";".join(sorted(set(to_list)))
        mail.CC = ";".join(sorted(set(cc_list)))

        # =========================
        # LOGO
        # =========================
        if os.path.exists(LOGO_PATH):
            logo = mail.Attachments.Add(LOGO_PATH)
            logo.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                "logoimg"
            )

        # =========================
        # EXCEL ATTACHMENT
        # =========================
        mail.Attachments.Add(file_path)

        # =========================
        # FINAL BODY
        # =========================
        mail.HTMLBody = body + build_signature() + mail.HTMLBody

        mail.Save()
        mail.Display(True)

        print("✅ Mail Created For:", wh)

    except Exception as e:
        print("❌ Error:", str(e))
        traceback.print_exc()

print("\n✅ ALL MAILS COMPLETED")