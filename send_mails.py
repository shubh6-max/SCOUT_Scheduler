import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

# === CONFIG ===
EXCEL_PATH = "leads_with_status.xlsx"
LOG_PATH = "email_log.csv"

SMTP_SERVER = "smtp.gmail.com"  # Use smtp.office365.com for Outlook
SMTP_PORT = 587
SENDER_EMAIL = "trailerbuddy96@gmail.com"
SENDER_PASSWORD = "whcy thim nsft kjcy"  # Use app password

BASE_FORM_URL = "http://192.168.0.103:8501"  # Replace with your deployed Streamlit form URL

def main():
    # --- Load Excel ---
    try:
        df = pd.read_excel(EXCEL_PATH)
    except Exception as e:
        print(f"❌ Failed to load Excel: {e}")
        return

    not_done_df = df[df["Status"].str.lower() == "not done"]

    # --- Group Leads by Stakeholder Email ---
    email_map = {}
    for _, row in not_done_df.iterrows():
        lead_name = row["Target Lead Name"]
        raw_emails = row["Leadership contact email"]
        for email in str(raw_emails).split(";"):
            email = email.strip().lower()
            if email not in email_map:
                email_map[email] = []
            email_map[email].append(lead_name)

    # --- Send Emails ---
    log_rows = []

    for stakeholder_email, leads in email_map.items():
        if not leads:
            continue

        form_link = f"{BASE_FORM_URL}/?email={stakeholder_email}"
        lead_list_html = "".join(f"<li>{lead}</li>" for lead in leads)

        html_body = f"""
        <html>
        <body>
            <p>Hi {stakeholder_email},</p>
            <p>Please help us by rating your relationship strength with the following leads:</p>
            <ul>{lead_list_html}</ul>
            <p>👉 <a href="{form_link}">Click here to open your form</a></p>
            <p>Thank you!</p>
        </body>
        </html>
        """

        msg = MIMEMultipart("alternative")
        msg["Subject"] = "📝 Warm Outreach - Your Action Needed"
        msg["From"] = SENDER_EMAIL
        msg["To"] = stakeholder_email
        msg.attach(MIMEText(html_body, "html"))

        try:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls()
                server.login(SENDER_EMAIL, SENDER_PASSWORD)
                server.sendmail(SENDER_EMAIL, stakeholder_email, msg.as_string())
                print(f"✅ Email sent to {stakeholder_email}")

            log_rows.append({
                "Email": stakeholder_email,
                "Num_Pending_Leads": len(leads),
                "Lead_Names": ", ".join(leads),
                "Form_Link": form_link,
                "Status": "Success",
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })

        except Exception as e:
            print(f"❌ Failed to send to {stakeholder_email}: {str(e)}")
            log_rows.append({
                "Email": stakeholder_email,
                "Num_Pending_Leads": len(leads),
                "Lead_Names": ", ".join(leads),
                "Form_Link": form_link,
                "Status": f"Failed: {str(e)}",
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })

    # --- Save Email Log ---
    pd.DataFrame(log_rows).to_csv(LOG_PATH, index=False)
    print(f"\n📄 Email log saved to {LOG_PATH}")

# === Run main if executed directly ===
if __name__ == "__main__":
    main()
