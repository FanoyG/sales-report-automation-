import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import os
from datetime import datetime

import smtplib
from email.message import EmailMessage

# 📌 Configurations
EXCEL_FILE = "Top_5_Selling_Products.xlsx"
GOOGLE_SHEET_NAME = "Multiple-sheet-automation"  # Change to your Google Sheet name
LOG_FILE = "upload_log.txt"
EMAIL_SENDER = "ismartk13@gmail.com"  # Change to your email
EMAIL_RECEIVER = "akstyle201471@gmail.com"  # Change to the recipient email
SERVICE_ACCOUNT_FILE = "multiple-sheet-automation-fb48f84f66a2.json"  # Replace with your credentials JSON file

# 🔗 Google Sheets API Authentication
scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
client = gspread.authorize(creds)

# 📊 Load Excel File
print("📥 Reading Excel file...")
try:
    df_summary = pd.read_excel(EXCEL_FILE, sheet_name="Summary")
    df_raw = pd.read_excel(EXCEL_FILE, sheet_name="Raw Data")
    # 🧹 Clean and prepare the data before uploading to Google Sheets
    df_summary.dropna(how="all", inplace=True)  # Remove fully empty rows
    df_raw.dropna(how="all", inplace=True)

    df_summary.fillna("", inplace=True)  # Replace NaN with empty string
    df_raw.fillna("", inplace=True)

    # Strip column names to avoid hidden spaces
    df_summary.columns = df_summary.columns.str.strip()
    df_raw.columns = df_raw.columns.str.strip()

except Exception as e:
    print(f"❌ Error reading Excel file: {e}")
    exit()

# 📤 Upload Data to Google Sheets
try:
    sheet = client.open(GOOGLE_SHEET_NAME)

    # Upload Summary Sheet
    worksheet_summary = sheet.worksheet("Summary") if "Summary" in [ws.title for ws in sheet.worksheets()] else sheet.add_worksheet(title="Summary", rows=100, cols=10)
    worksheet_summary.clear()
    worksheet_summary.update([df_summary.columns.values.tolist()] + df_summary.values.tolist())

    # Upload Raw Data Sheet
    worksheet_raw = sheet.worksheet("Raw Data") if "Raw Data" in [ws.title for ws in sheet.worksheets()] else sheet.add_worksheet(title="Raw Data", rows=100, cols=10)
    worksheet_raw.clear()
    worksheet_raw.update([df_raw.columns.values.tolist()] + df_raw.values.tolist())

    print("✅ Data uploaded successfully!")
    log_message = "✅ Data successfully uploaded to Google Sheets.\n"

except Exception as e:
    log_message = f"❌ Error uploading data to Google Sheets: {e}\n"
    print(log_message)

# 📅 Get current date and time
now = datetime.now()
date_str = now.strftime("%Y-%m-%d")
time_str = now.strftime("%H:%M:%S")

# 📝 Define log header with Process 1
log_header = f"""
==================================================
📅 Date: {date_str}   ⏰ Time: {time_str}
==================================================
🚀 Process 1: Data Upload Started
--------------------------------------------------
"""

# 📝 Write Log File
with open(LOG_FILE, "a", encoding="utf-8") as log:
    log.write(log_header)

# 📩 Send Email Notification
def send_email():
    try:
        msg = EmailMessage()
        msg["Subject"] = "📊 Sales Report Updated!"
        msg["From"] = EMAIL_SENDER
        msg["To"] = EMAIL_RECEIVER
        msg.set_content(f"The sales report has been successfully updated.\n\nView it here: https://docs.google.com/spreadsheets/d/{sheet.id}")

        # SMTP Connection
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(EMAIL_SENDER, "ffak dfvp djrj dntl")  # Use App Password for security
            smtp.send_message(msg)

        print("✅ Email sent successfully!")
        with open(LOG_FILE, "a", encoding="utf-8") as log:
            log.write("✅ Email notification sent.\n")

    except Exception as e:
        print(f"❌ Error sending email: {e}")
        with open(LOG_FILE, "a", encoding="utf-8") as log:
            log.write(f"❌ Error sending email: {e}\n")

send_email()
