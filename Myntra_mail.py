
import pandas as pd
import win32com.client as win32
import os

# =========================
# SETTINGS
# =========================
FILE_PATH = r"C:\Users\Dilip\Downloads\Myntra_output_17-06-2026_16-32.xlsx"

LOGO_PATH = r"C:\Users\Dilip\Downloads\Outlook-5uukthm0.png"

MARKETPLACE = "MYNTRA"

# =========================
# READ EXCEL
# =========================
df = pd.read_excel(FILE_PATH)

df.columns = df.columns.astype(str).str.strip()

# =========================
# FILTER NOT FOUND
# =========================
df["Found_In_Sheet"] = (
    df["Found_In_Sheet"]
    .astype(str)
    .str.upper()
    .str.strip()
)

df = df[
    df["Found_In_Sheet"] == "NOT FOUND"
]

# =========================
# REQUIRED COLUMNS
# =========================
required_cols = [
    "warehouse_id",
    "order_id",
    "seller_order_id",
    "forward_tracking_number_1",
    "forward_tracking_number_2",
    "return_tracking_number",
    "deliver_to_seller_date"
]

df = df[required_cols]

# =========================
# EMAIL MAPPING
# =========================
warehouse_email_map = {

    "BWD": {
        "TO": [
            "sameer.bhoir@maersk.com",
            "nilesh.bajage@apmterminals.com"
        ],
        "CC": [
            "chetan@happyecom.com"
        ]
    },

    "KOL": {
        "TO": [
            "tutu.hazra@apmterminals.com"
        ],
        "CC": [
            "tuhin@happyecom.com"
        ]
    },

    "GGN": {
        "TO": [
            "ramesh.kumar@apmterminals.com",
            "nitin.sharma@apmterminals.com"
        ],
        "CC": [
            "alok@happyecom.com",
            "naven.singh@maersk.com"
        ]
    },

    "BNG": {
        "TO": [
            "padmavathamma.sr3@flipkart.com",
            "basvaraj@happyecom.com"
        ],
        "CC": [
            "sivakumar.v1@flipkart.com"
        ]
    },

    "AHMD": {
        "TO": [
            "ops.ahmedabad@emizainc.com"
        ],
        "CC": []
    },

    "ASSAM": {
        "TO": [
            "debarup.chanda@apmterminals.com",
            "chetan@happyecom.com",
            "tutu.hazra@apmterminals.com"
        ],
        "CC": [
            "Sumit.boro@apmterminals.com",
            "ajeet.kumar.sharma@maersk.com",
            "ravi.ranjan.4@maersk.com"
        ]
    }
}

# =========================
# COMMON CC
# =========================
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
# GROUP BY WAREHOUSE
# =========================
grouped = df.groupby("warehouse_id")

# =========================
# PROCESS
# =========================
for wh, data in grouped:

    wh = str(wh).strip().upper()

    print(f"\nProcessing Warehouse -> {wh}")

    # =========================
    # EMAIL MAP CHECK
    # =========================
    if wh not in warehouse_email_map:

        print("❌ No email mapping found")

        continue

    # =========================
    # CREATE EXCEL FILE
    # =========================
    file_name = f"{MARKETPLACE}_{wh}_NOT_FOUND.xlsx"

    file_path = os.path.join(
        os.getcwd(),
        file_name
    )

    data.to_excel(file_path, index=False)

    # =========================
    # SUBJECT
    # =========================
    subject_line = (
        f"{MARKETPLACE} | {wh} | "
        f"Shipment Marked Delivered but Not Received at Warehouse"
    )

    # =========================
    # TRAIL SEARCH
    # =========================
    mail = None

    messages = sent_folder.Items.Restrict(
        "[MessageClass] = 'IPM.Note'"
    )

    messages.Sort("[SentOn]", True)

    for msg in messages:

        try:

            subject = str(msg.Subject).strip().lower()

            # OLD SUBJECT
            old_subject = (
                f"{wh} | "
                f"Shipment Marked Delivered but Not Received at Warehouse"
            ).lower()

            # NEW SUBJECT
            new_subject = (
                f"{MARKETPLACE} | {wh} | "
                f"Shipment Marked Delivered but Not Received at Warehouse"
            ).lower()

            # EXACT MATCH ONLY
            if (
                subject == old_subject
                or subject == new_subject
            ):

                mail = msg.ReplyAll()

                print("✅ Trail Mail Found")

                break

        except:
            pass

    # =========================
    # NEW MAIL
    # =========================
    if mail is None:

        print("⚠️ New Mail Created")

        mail = outlook.CreateItem(0)

        mail.Subject = subject_line

    # =========================
    # BODY
    # =========================
    body = f"""
    <div style="font-family:Calibri;font-size:11pt;">

    Hi Team,<br><br>

    We have observed that the below shipments are marked as delivered on the marketplace;
    however, the same has not been received at the warehouse.<br><br>

    Request you to kindly check and confirm the status from your end.<br><br>

    <b>Total Cases: {len(data)}</b><br><br>

    </div>
    """

    # =========================
    # EMAILS
    # =========================
    to_list = warehouse_email_map[wh]["TO"]

    cc_list = (
        warehouse_email_map[wh]["CC"]
        + COMMON_CC
    )

    # =========================
    # BNG SPECIAL
    # =========================
    if wh == "BNG":

        to_list.append(
            "bala.murali@flipkart.com"
        )

        cc_list.append(
            "syed.yousuf@flipkart.com"
        )

    # =========================
    # REMOVE DUPLICATES
    # =========================
    mail.To = ";".join(set(to_list))

    mail.CC = ";".join(set(cc_list))

    # =========================
    # ATTACH LOGO
    # =========================
    logo = mail.Attachments.Add(LOGO_PATH)

    logo.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
        "logoimg"
    )

    # =========================
    # ATTACH EXCEL
    # =========================
    mail.Attachments.Add(file_path)

    # =========================
    # FINAL BODY
    # =========================
    mail.HTMLBody = (
        body
        + build_signature()
        + mail.HTMLBody
    )

    # =========================
    # DISPLAY MAIL
    # =========================
    mail.Display()

print("\n✅ ALL MYNTRA MAILS CREATED SUCCESSFULLY")