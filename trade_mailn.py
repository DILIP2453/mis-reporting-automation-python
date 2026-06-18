import pandas as pd
import win32com.client as win32
import os

# =========================
# SETTINGS
# =========================

TRAIL_MODE = True

DATA_FILE = r"C:\Users\Dilip\Downloads\RETURN LSD 16-06-2026.xlsx"
LOGO_PATH = r"C:\Users\Dilip\Downloads\Outlook-5uukthm0.png"

# =========================
# READ EXCEL
# =========================

df = pd.read_excel(DATA_FILE, sheet_name="Sheet1")

# CLEAN COLUMN NAMES
df.columns = (
    df.columns.astype(str)
    .str.strip()
    .str.replace(r"\s+", " ", regex=True)
)

received_col = "Received Location"
remark_col = "Remarks"

# CLEAN VALUES
df[received_col] = (
    df[received_col]
    .astype(str)
    .str.strip()
    .str.upper()
)

df[remark_col] = (
    df[remark_col]
    .astype(str)
    .str.strip()
    .str.upper()
)

# REMOVE BLANKS
df = df[df[received_col] != ""]
df = df[df[remark_col] != "NOT PRESENT IN SALES TABLE"]

# =========================
# LOCATION MAP
# =========================

location_short_map = {
    "BWD": "LSD-MAERSK-BWD",
    "KOL": "LSD-MAERSK-KOL",
    "GGN": "LSD-MAERSK-GGN",
    "BNG": "LSD-EKART-BNG",
    "AHMD": "LSD-EMIZA-AHMD",
    "ASSAM": "LSD-MAERSK-ASSAM",
    "CHN": "LSD-EMIZA-CHENNAI",
    "HYD": "LSD-EMIZA-HYD",
    "LKN": "LSD-EMIZA-LKN"
}

# =========================
# REQUIRED COLUMNS
# =========================

required_columns = [
    "ID",
    "Tracking ID",
    "DATE",
    "Received By",
    "Received Location",
    "Order ID",
    "MARKET PLACE ID",
    "Vinculum Order ID",
    "Channel Type",
    "Remarks"
]

df = df[required_columns]

# =========================
# EMAIL MAP
# =========================

location_emails = {

    "LSD-MAERSK-BWD": {
        "TO": [
            "sameer.bhoir@maersk.com",
            "nilesh.bajage@apmterminals.com"
        ],
        "CC": [
            "chetan@happyecom.com"
        ]
    },

    "LSD-MAERSK-KOL": {
        "TO": [
            "tutu.hazra@apmterminals.com"
        ],
        "CC": [
            "tuhin@happyecom.com"
        ]
    },

    "LSD-MAERSK-GGN": {
        "TO": [
            "ramesh.kumar@apmterminals.com",
            "nitin.sharma@apmterminals.com"
        ],
        "CC": [
            "alok@happyecom.com",
            "naven.singh@maersk.com"
        ]
    },

    "LSD-EKART-BNG": {
        "TO": [
            "padmavathamma.sr3@flipkart.com",
            "basvaraj@happyecom.com"
        ],
        "CC": [
            "sivakumar.v1@flipkart.com"
        ]
    },

    "LSD-EMIZA-AHMD": {
        "TO": [
            "ops.ahmedabad@emizainc.com",
            "chetan@happyecom.com"
        ],
        "CC": []
    },

    "LSD-MAERSK-ASSAM": {
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
    
"LSD-EMIZA-CHENNAI": {
    "TO": [
        "manoj.s@emizainc.net",
        "boopathy.d@emizainc.com"
    ],
    "CC": [
        "hariprakash.r@emizainc.com"
    ]
},


"LSD-EMIZA-HYD": {
    "TO": [
        "v.teja@emizainc.com"
    ],
    "CC": []
},
 
   "LSD-EMIZA-LKN": {
      "TO": [
        "sachin.yadav@emizainc.net", 
         "ops.lucknow@emizainc.com"
    ],
    "CC": [
          "alok@happyecom.com",
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
# OUTLOOK SETUP
# =========================

outlook = win32.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# SENT ITEMS
sent_folder = namespace.GetDefaultFolder(5)

# =========================
# SIGNATURE
# =========================

def build_signature():
    return """
    <br>

    <b style="color:#F37021;font-size:14px;">
    Dilip Patekar
    </b><br><br>

    <b style="color:#F37021;">HAPPILY</b><br>

    MIS Associate<br><br>

    Phone: +91 8169881923<br>

    <img src="cid:logoimg" width="120"><br><br>
    """

# =========================
# MAIN PROCESS
# =========================

grouped = df.groupby(received_col)

for location, location_data in grouped:

    mapped_location = location_short_map.get(location, location)

    print(f"\nProcessing -> {location}")
    print(f"Mapped -> {mapped_location}")

    # =========================
    # SKIP IF LOCATION NOT FOUND
    # =========================

    if mapped_location not in location_emails:

        print("❌ Location Not Mapped → Skipped")

        continue

    # =========================
    # CREATE EXCEL FILE
    # =========================

    file_name = f"{mapped_location}_MAPPING STATUS.xlsx"

    file_path = os.path.join(os.getcwd(), file_name)

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:

        for remark, remark_data in location_data.groupby(remark_col):

            export_columns = [
                "ID",
                "Tracking ID",
                "DATE",
                "Received By",
                "Received Location",
                "Order ID",
                "MARKET PLACE ID",
                "Vinculum Order ID",
                "Channel Type"
            ]

            remark_data[export_columns].to_excel(
                writer,
                sheet_name=remark[:30],
                index=False
            )

    print(f"✅ Excel Created -> {file_name}")

    # =========================
    # SUMMARY
    # =========================

    summary_counts = location_data[remark_col].value_counts()

    total_cases = len(location_data)

    summary_html = ""

    for remark, count in summary_counts.items():

        summary_html += f"• {remark.title()} – {count}<br>"

    summary_html += f"""
    <br>
    <b>Total Cases – {total_cases}</b>
    <br><br>
    """

    # =========================
    # EMAIL BODY
    # =========================

    body = f"""
    <div style="font-family:Calibri;font-size:11pt;">

    Hi Team,<br><br>

    Kindly refer to the attached sheet.<br><br>

    <b>Summary:</b><br>

    {summary_html}

    </div>
    """

    # =========================
    # EMAIL RECIPIENTS
    # =========================

    expected_to = location_emails[mapped_location]["TO"].copy()

    expected_cc = (
        location_emails[mapped_location]["CC"].copy()
        + COMMON_CC
    )

    # BNG SPECIAL CASE
    if mapped_location == "LSD-EKART-BNG":

        expected_to.append("bala.murali@flipkart.com")

        expected_cc.append("syed.yousuf@flipkart.com")

    # REMOVE DUPLICATES
    final_to = list(set(expected_to))

    final_cc = list(set(expected_cc))

    # =========================
    # SUBJECT
    # =========================

    subject_line = f"HEC Returns Mapping Status ({mapped_location})"

    # =========================
    # FIND LATEST EXACT TRAIL MAIL
    # =========================

    mail = None

    latest_msg = None

    messages = sent_folder.Items

    expected_subject = subject_line.strip().lower()

    print(f"🔍 Searching Trail -> {subject_line}")

    for msg in messages:

        try:

            # ONLY MAIL ITEMS
            if msg.Class != 43:
                continue

            subject = str(msg.Subject).strip().lower()

            # REMOVE RE / FW PREFIX
            clean_subject = (
                subject.replace("re:", "")
                       .replace("fw:", "")
                       .strip()
            )

            # EXACT SUBJECT MATCH
            if clean_subject == expected_subject:

                # PICK LATEST MAIL
                if latest_msg is None:

                    latest_msg = msg

                else:

                    if msg.SentOn > latest_msg.SentOn:

                        latest_msg = msg

        except Exception as e:

            print(f"Error Reading Mail: {e}")

    # =========================
    # CREATE REPLY FROM LATEST MAIL
    # =========================

    if latest_msg is not None:

        print("✅ Latest Trail Mail Found")

        print(f"📩 Subject -> {latest_msg.Subject}")

        print(f"📅 Sent On -> {latest_msg.SentOn}")

        mail = latest_msg.ReplyAll()

    else:

        print("⚠️ No Trail Found → Creating New Mail")

        mail = outlook.CreateItem(0)

        mail.Subject = subject_line

    # =========================
    # FORCE TO / CC
    # =========================

    mail.To = ";".join(final_to)

    mail.CC = ";".join(final_cc)

    # =========================
    # ATTACH LOGO
    # =========================

    logo_attachment = mail.Attachments.Add(LOGO_PATH)

    logo_attachment.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
        "logoimg"
    )

    # =========================
    # ATTACH EXCEL FILE
    # =========================

    mail.Attachments.Add(file_path)

    # =========================
    # HTML BODY
    # =========================

    mail.HTMLBody = (
        body
        + build_signature()
        + mail.HTMLBody
    )

    # =========================
    # DISPLAY / SEND
    # =========================

    if TRAIL_MODE:

        mail.Display()

    else:

        mail.Send()

        print("✅ Mail Sent")

print("\n✅ Process Completed Successfully")