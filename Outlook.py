import streamlit as st
import pandas as pd
import smtplib
import time
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os

# Page config and title
st.set_page_config(page_title="Simple Outlook Email Sender", layout="wide")
st.title("📧 Simple Outlook Email Sender")

# Sidebar: Credentials & settings
with st.sidebar:
    st.header("⚙️ Email Settings")

    username = st.text_input("Outlook Email", placeholder="yourname@outlook.com")
    password = st.text_input("Outlook Password", type="password")

    enable_cc = st.checkbox("Enable CC Email")
    if enable_cc:
        cc_email = st.text_input("CC Email", placeholder="cc@example.com")
    else:
        cc_email = ""

    delay = st.slider("Delay between emails (seconds)", 1, 60, 5)
    read_receipt = st.checkbox("Request Read Receipt")

# Sample Excel download
if st.button("📥 Download Sample Excel File"):
    df_sample = pd.DataFrame({'Name': ['John Doe'], 'Email': ['john@example.com']})
    df_sample.to_excel("sample_email_list.xlsx", index=False)
    with open("sample_email_list.xlsx", "rb") as f:
        st.download_button("Download Sample Excel", f, file_name="sample_email_list.xlsx")

# Upload Excel file
email_file = st.file_uploader("Upload Excel file with 'Name' and 'Email'", type=["xlsx"])

# Email template input
st.subheader("✍️ Email Template")
st.markdown("""
Write your email template here. Use `{Name}` to insert the recipient's first name.
email_template = st.text_area("Email Template", height=200, value="Hello {Name},\n\nThis is a test email.\n\nBest regards,\nYour Name")

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
        st.error(f"Failed to send email to {to_email}: {e}")
        return False

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

if st.button("🚀 Start Sending Emails"):
    if not username or not password:
        st.error("Please enter your Outlook email and password.")
        st.stop()
    if not email_file:
        st.error("Please upload your Excel file.")
        st.stop()
    if not email_template.strip():
        st.error("Please enter an email template.")
        st.stop()

    try:
        df = pd.read_excel(email_file)
        df.columns = df.columns.str.strip()
        if 'Name' not in df.columns or 'Email' not in df.columns:
            st.error("Excel file must contain 'Name' and 'Email' columns.")
            st.stop()

        # Track progress - emails already sent
        progress_path = 'sent_emails.csv'
        sent_emails = set()
        if os.path.exists(progress_path):
            sent_emails = set(pd.read_csv(progress_path)['Email'].tolist())

        df = df[~df['Email'].isin(sent_emails)].reset_index(drop=True)

        failed = []
        total = len(df)

        for i, row in df.iterrows():
            full_name = row['Name']
            first_name = full_name.split()[0] if isinstance(full_name, str) else ""
            email = row['Email']

            subject = f"Hello {full_name}, Important Information"
            body = email_template.replace("{Name}", first_name)
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            with st.spinner(f"Sending email {i+1}/{total} to {email}..."):
                success = send_email(subject, body, email, cc_email, read_receipt)
                status = "Sent" if success else "Failed"
                log_email(full_name, email, subject, status, timestamp)

                if success:
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

# Email log dashboard
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
