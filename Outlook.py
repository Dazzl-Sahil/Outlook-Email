import streamlit as st
import pandas as pd
import smtplib
import time
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os

# ----- Load Outlook credentials from Streamlit secrets -----
username = st.secrets["outlook_email"]
password = st.secrets["outlook_password"]

# ----- Function to send email -----
def send_email(subject, body, to_email, cc_email="", read_receipt=False):
    msg = MIMEMultipart()
    msg['From'] = username
    msg['To'] = to_email
    if cc_email:
        msg['CC'] = cc_email
    msg['Subject'] = subject

    if read_receipt:
        msg.add_header('Disposition-Notification-To', username)

    msg.attach(MIMEText(body, 'plain'))

    recipients = [to_email]
    if cc_email:
        recipients.append(cc_email)

    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(username, password)
        server.sendmail(username, recipients, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")
        return False

# ----- Email logging -----
def log_email(name, email, subject, status, timestamp):
    log_file = 'email_log.csv'
    entry = pd.DataFrame([{
        "Timestamp": timestamp,
        "Name": name,
        "Email": email,
        "Subject": subject,
        "Status": status
    }])
    if os.path.exists(log_file):
        existing = pd.read_csv(log_file)
        updated = pd.concat([existing, entry], ignore_index=True)
    else:
        updated = entry
    updated.to_csv(log_file, index=False)

# ----- Main UI -----
st.set_page_config(page_title="Outlook Email Sender", layout="wide")
st.title("📧 Automated Outlook Email Sender")

# Sample file download button
if st.button("📥 Download Sample Excel"):
    df_sample = pd.DataFrame({'Name': ['John Doe'], 'Email': ['john@example.com']})
    df_sample.to_excel("sample_email_list.xlsx", index=False)
    with open("sample_email_list.xlsx", "rb") as f:
        st.download_button("Download Sample Excel", f, file_name="sample_email_list.xlsx")

# File uploads
email_file = st.file_uploader("📄 Upload Excel File with 'Name' and 'Email' columns", type=["xlsx"])
template_file = st.file_uploader("📝 Upload Email Template (.txt file)", type=["txt"])

# Delay slider
delay = st.slider("⏱️ Delay Between Emails (seconds)", 1, 60, 5)

# Read receipt toggle
read_receipt = st.checkbox("📩 Request Read Receipt")

# CC enable toggle + CC email input
enable_cc = st.checkbox("Enable CC Email")
if enable_cc:
    cc_email = st.text_input("CC Email", value=username, help="Enter email to CC on all emails.")
else:
    cc_email = ""

# Start button
if st.button("🚀 Start Sending Emails"):

    if not email_file or not template_file:
        st.error("Please upload both the Excel file and the email template.")
        st.stop()

    try:
        df = pd.read_excel(email_file)
        df.columns = df.columns.str.strip()

        if 'Name' not in df.columns or 'Email' not in df.columns:
            st.error("The Excel file must contain 'Name' and 'Email' columns.")
            st.stop()

        template = template_file.read().decode("utf-8")

        # Load progress of sent emails (to resume)
        progress_path = 'sent_emails.csv'
        sent_emails = set()
        if os.path.exists(progress_path):
            sent_emails = set(pd.read_csv(progress_path)['Email'].tolist())

        # Filter out already sent emails
        df = df[~df['Email'].isin(sent_emails)].reset_index(drop=True)

        failed = []
        total = len(df)

        for i, row in df.iterrows():
            full_name = row['Name']
            first_name = full_name.split()[0] if isinstance(full_name, str) else ""
            email = row['Email']

            subject = f"Hello {full_name}, Important Information"
            body = template.replace("{Name}", first_name)
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            with st.spinner(f"Sending email {i+1}/{total} to {email}..."):
                success = send_email(subject, body, email, cc_email, read_receipt)
                status = "Sent" if success else "Failed"
                log_email(full_name, email, subject, status, timestamp)

                if success:
                    # Append to sent_emails.csv
                    with open(progress_path, "a") as f:
                        f.write(f"{email}\n")
                else:
                    failed.append({"Name": full_name, "Email": email})

                time.sleep(delay)

        if failed:
            fail_df = pd.DataFrame(failed)
            fail_df.to_excel("failed_emails.xlsx", index=False)
            with open("failed_emails.xlsx", "rb") as f:
                st.download_button("⬇️ Download Failed Emails", f, file_name="failed_emails.xlsx")
            st.warning(f"{len(failed)} emails failed.")
        else:
            st.success("✅ All emails sent successfully!")

    except Exception as e:
        st.error(f"Error: {e}")

# ----- Email log dashboard -----
st.markdown("---")
st.subheader("📊 Email Log Dashboard")

if os.path.exists("email_log.csv"):
    log_df = pd.read_csv("email_log.csv")
    st.dataframe(log_df, use_container_width=True)
    st.download_button("⬇️ Download Email Log CSV", log_df.to_csv(index=False).encode('utf-8'), "email_log.csv")

    if st.button("🗑️ Clear Email Log"):
        os.remove("email_log.csv")
        st.success("Email log cleared.")
else:
    st.info("No email log yet.")
