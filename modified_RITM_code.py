import win32com.client
import pandas as pd
import os
import time
import datetime
import pywhatkit
import requests

# ---------------- SETTINGS ----------------
BASE_PATH = r"C:\Automation_mail_RITM"
RITM_FILE = os.path.join(BASE_PATH, "RITM_input.xlsx")       # RITM requests raised
REPORT_FILE = os.path.join(BASE_PATH, "RITM_Report.xlsx")    # Report with full data
OUTPUT_FILE = os.path.join(BASE_PATH, "Final_Output.xlsx")   # Optional attachment

COLUMNS_TO_PICK = [
    "Activity Details",
    "Access Type",
    "Node Name",
    "Access Required till Date",
    "Status",
    "Activity",
    "Remarks"
]

# WhatsApp Groups
N8_NUMBERS = ["+916390816637"]
E15_NUMBERS = ["+917777777777", "+916666666666"]
ENM_EXTRA_NUMBER = "+915555555555"

N8_CIRCLES = ["MH", "MU", "KO", "BH", "OR", "WB", "MP", "GUJ"]
E15_CIRCLES = ["AP", "KK", "KL", "TN", "CHN", "PB", "HR", "HP", "NE", "AS", "JK", "UPE", "UPW", "DL", "RJ"]

# Push notification settings
PUSHOVER_USER_KEY = "uopga7mwkot4a2de4a2zn8ipm87b8n"
PUSHOVER_API_TOKEN = "atkdhq3pp28qtzzcwoh9h4e6rjyxy4"

def send_push_notification(message):
    try:
        requests.post("https://api.pushover.net/1/messages.json", data={
            "token": PUSHOVER_API_TOKEN,
            "user": PUSHOVER_USER_KEY,
            "message": message
        })
    except Exception as e:
        print("Notification error:", e)

# ---------------- STEP 1: SAFE READ EXCEL FILES ----------------
# Read RITM requests
if os.path.exists(RITM_FILE):
    try:
        ritm_df = pd.read_excel(RITM_FILE)
        print("✅ RITM input loaded")
    except Exception as e:
        print("❌ Error reading RITM input:", e)
        ritm_df = pd.DataFrame(columns=["CIRCLE", "Executor", "Activity Details"])
else:
    print("⚠ RITM input file not found, creating empty table")
    ritm_df = pd.DataFrame(columns=["CIRCLE", "Executor", "Activity Details"])

# Read Report
if os.path.exists(REPORT_FILE):
    try:
        report_df = pd.read_excel(REPORT_FILE)
        print("✅ Report file loaded")
    except Exception as e:
        print("❌ Error reading Report file:", e)
        report_df = pd.DataFrame(columns=COLUMNS_TO_PICK)
else:
    print("⚠ Report file not found, creating empty report table")
    report_df = pd.DataFrame(columns=COLUMNS_TO_PICK)

ritm_df.columns = ritm_df.columns.str.strip()
report_df.columns = report_df.columns.str.strip()

# ---------------- STEP 2: FILTER REPORT BASED ON RITM REQUEST ----------------
filtered_report_rows = []

for _, ritm_row in ritm_df.iterrows():
    circle = str(ritm_row.get("CIRCLE", "")).upper()
    executor = str(ritm_row.get("Executor", ""))

    # Filter report: matching circle and executor/activity
    if "Activity Details" in report_df.columns:
        temp_df = report_df[
            (report_df["Activity Details"].str.contains(executor, case=False, na=False)) |
            (report_df["Activity Details"].str.contains(circle, case=False, na=False))
        ]
    else:
        temp_df = pd.DataFrame(columns=COLUMNS_TO_PICK)

    if temp_df.empty:
        filtered_report_rows.append({col: "No data available" for col in COLUMNS_TO_PICK})
    else:
        for _, r in temp_df.iterrows():
            filtered_report_rows.append({col: r[col] if col in r else "" for col in COLUMNS_TO_PICK})

final_df = pd.DataFrame(filtered_report_rows)

# Save output as optional attachment
try:
    final_df.to_excel(OUTPUT_FILE, index=False)
except Exception as e:
    print("⚠ Could not save attachment:", e)

# ---------------- STEP 3: CONVERT TO HTML TABLE ----------------
def df_to_html(df):
    header_style = 'style="background-color:#C6EFCE; border:1px solid black; text-align:center;"'
    cell_style = 'style="border:1px solid black; text-align:center;"'
    
    html = '<table border="0" cellspacing="0" cellpadding="5">'
    html += '<tr>'
    for col in df.columns:
        html += f'<th {header_style}>{col}</th>'
    html += '</tr>'
    
    for _, row in df.iterrows():
        html += '<tr>'
        for val in row:
            html += f'<td {cell_style}>{val}</td>'
        html += '</tr>'
    
    html += '</table>'
    return html

html_table = df_to_html(final_df)

# ---------------- STEP 4: SEND OUTLOOK EMAIL ----------------
outlook_app = win32com.client.Dispatch("Outlook.Application")
mail = outlook_app.CreateItem(0)

mail.Subject = "Root //ts_ user//admin access required for activities execution"
mail.To = "raghvendra.b.pratap.singh@ericsson.com"
mail.Importance = 2  # High importance

mail.HTMLBody = f"""
<p>@Deepan Kansal /Navneet Kumar Gupta Sir: Pls approve.</p>
<p>@Gagan Pruthi Sir: Pls approve Catalogue Request for E15 Circles Before 6:00 PM.</p>
<p>@Prabhat Kumar Prabhakar: Pls approve Catalogue Request for E15 Circle ENM’s before 6:00 PM.</p>
<p>@Tarun Kumar H Sir: Pls approve Catalogue Request for N8 Circles Before 6:00 PM.</p>

<p>Please find the RITM details below:</p>
{html_table}

<p>Regards,<br>Raghvendra</p>
"""

# Attach Excel file if exists
if os.path.exists(OUTPUT_FILE):
    try:
        mail.Attachments.Add(OUTPUT_FILE)
    except Exception as e:
        print("⚠ Could not attach file:", e)

mail.Send()
print("✅ Outlook email sent successfully.")

# ---------------- STEP 5: PREPARE RITM GROUPS FOR WHATSAPP ----------------
n8_ritms = []
e15_ritms = []
enm_ritms = []

for _, row in ritm_df.iterrows():
    circle = str(row.get("CIRCLE", "")).upper()
    ritm_no = str(row.get("Activity Details", ""))

    if "ENM" in circle:
        enm_ritms.append(ritm_no)
    elif circle in N8_CIRCLES:
        n8_ritms.append(ritm_no)
    elif circle in E15_CIRCLES:
        e15_ritms.append(ritm_no)

# ---------------- STEP 6: CLOSE CHROME ----------------
os.system("taskkill /f /im chrome.exe")
time.sleep(5)

# ---------------- STEP 7: SEND WHATSAPP ----------------
now = datetime.datetime.now()
hour = now.hour
minute = now.minute + 2

def send_message(numbers, message):
    global hour, minute
    for num in numbers:
        try:
            pywhatkit.sendwhatmsg(num, message, hour, minute, wait_time=10, tab_close=True)
            minute += 1
        except Exception as e:
            print("Error sending to", num, e)

if n8_ritms:
    send_message(N8_NUMBERS, "Hi Sir,\n\nPlease approve the RITM requests.\n\nN8 RITMs:\n" + "\n".join(n8_ritms))
if e15_ritms:
    send_message(E15_NUMBERS, "Hi Sir,\n\nPlease approve the RITM requests.\n\nE15 RITMs:\n" + "\n".join(e15_ritms))
if enm_ritms:
    send_message([ENM_EXTRA_NUMBER], "Hi Sir,\n\nPlease approve the RITM requests.\n\nENM RITMs:\n" + "\n".join(enm_ritms))

# ---------------- STEP 8: PUSH NOTIFICATION ----------------
send_push_notification("✅ WhatsApp messages sent successfully!")

print("✅ ALL TASK COMPLETED")