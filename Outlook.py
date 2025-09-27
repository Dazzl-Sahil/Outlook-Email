import streamlit as st
import pandas as pd
import smtplib
import time
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os

# ------------------ UI Configuration ------------------
st.set_page_config(page_title="Universal Email Sender", layout="wide")
st.title("📨 Universal Email Sender (Gmail or Outlook)")

# ------------------ Sidebar Settings ------------------
with st.sidebar:
    st.header("📧 Email Settings")
    
    provider = st.selectbox("Choose Email Provider", ["Gmail", "Outlook"])
    username = st.text_input("Email Address", placeholder="yourname@gmail.com or yourname@outlook.com")
    password = st.text_input("App Password", type="password", placeholder="Paste your App Password")

    st.markdown("---")
    delay = st.slider("Delay between emails (in seconds)", 1, 60, 5)

    read_receipt = st.checkbox("Request Read Receipt")

    enable_cc = st.checkbox("Enable CC Email")
    cc_email = st.text_input("CC Email", placeholder="someone@domain.com") if enable_cc else ""

# ------------------ Sample File Download ------------------
if st.button("📥 Download Sample Excel"):
    df_sample = pd.DataFrame({'Name': ['John Doe'], 'Email': ['john@example.com']})
    df_sample.to_excel("sample_email_list.xlsx", index=False)
    with open("sample_email_list.xlsx", "rb") as f:
        st.download_button("Download Sample Excel", f, file_name="sample_email_list.xlsx")

# ------------------ Upload Excel ------------------
email_file = st.file_uploader("Upload Excel with columns 'Name' and 'Email'", type=["xlsx"])

df = None
if email_file:
    try:
        df = pd.read_excel(email_file)
        df.columns = df.columns.str.strip()
        if 'Name' not in df.columns or 'Email' not in df.columns:
            st.error("❌ Excel must contain 'Name' and 'Email' columns.")
            df = None
        else:
            st.success("✅ Excel loaded successfully!")
            st.subheader("👥 Preview of Recipients")
            st.dataframe(df[['Name', 'Email']], use_container_width=True)
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")

# ------------------ Email Template ------------------
st.subheader("✍️ Email Template")
email_template = st.text_area(
    "Write your email. Use `{Name}` to insert first name.",
    height=200,
    value="Hello {Name},\n\nThis is a test email.\n\nBest regards,\nYour Name"
)

# ------------------ SMTP Send Function ------------------
def get_smtp_settings(provider):
    if provider == "Gmail":
        return "smtp.gmail.com", 587
    elif provider == "Outlook":
        return "smtp.office365.com", 587
    else:
        raise ValueError("Unsupported email provider")

def send_email(subject, body, to_email, cc_email="", read_receipt=False):
    msg = MIMEMultipart()
    msg["From"] = username
    msg["To"] = to_email
    msg["Subject"] = subject
    if cc_email:
        msg["CC"] = cc_email
    if read_receipt:
        msg.add_header("Disposition-Notification-To", username)

    msg.attach(MIMEText(body, "plain"))

    recipients = [to_email] + ([cc_email] if cc_email else [])

    smtp_server, port = get_smtp_settings(provider)

    try:
        server = smtplib.SMTP(smtp_server, port)
        server.starttls()
        server.login(username, password)
        server.sendmail(username, recipients, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"❌ Failed to send email to {to_email}: {e}")
        return False

# ------------------ Log Email ------------------
def log_email(name, email, subject, status, timestamp):
    log_file = "email_log.csv"
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

# ------------------ Start Sending Emails ------------------
if st.button("🚀 Start Sending Emails"):
    if not username or not password:
        st.error("Enter your email and App Password in the sidebar.")
        st.stop()
    if df is None or df.empty:
        st.error("Upload a valid Excel file with 'Name' and 'Email'.")
        st.stop()
    if not email_template.strip():
        st.error("Email template cannot be empty.")
        st.stop()

    progress_file = "sent_emails.csv"
    sent_emails = set()
    if os.path.exists(progress_file):
        sent_emails = set(pd.read_csv(progress_file)["Email"].tolist())

    df_to_send = df[~df["Email"].isin(sent_emails)].reset_index(drop=True)

    failed = []
    total = len(df_to_send)

    for i, row in df_to_send.iterrows():
        full_name = row["Name"]
        first_name = full_name.split()[0] if isinstance(full_name, str) else ""
        to_email = row["Email"]

        subject = f"Hello {full_name}, Important Information"
        body = email_template.replace("{Name}", first_name)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with st.spinner(f"Sending email {i+1}/{total} to {to_email}..."):
            success = send_email(subject, body, to_email, cc_email, read_receipt)
            status = "Sent" if success else "Failed"
            log_email(full_name, to_email, subject, status, timestamp)

            if success:
                with open(progress_file, "a") as f:
                    f.write(f"{to_email}\n")
            else:
                failed.append({"Name": full_name, "Email": to_email})

            time.sleep(delay)

    # Download failed
    if failed:
        fail_df = pd.DataFrame(failed)
        fail_df.to_excel("failed_emails.xlsx", index=False)
        with open("failed_emails.xlsx", "rb") as f:
            st.download_button("⬇️ Download Failed Emails", f, file_name="failed_emails.xlsx")
        st.warning(f"⚠️ {len(failed)} emails failed to send.")
    else:
        st.success("🎉 All emails sent successfully!")

# ------------------ Dashboard ------------------
st.markdown("---")
st.subheader("📊 Email Log Dashboard")

if os.path.exists("email_log.csv"):
    log_df = pd.read_csv("email_log.csv")
    st.dataframe(log_df, use_container_width=True)
    st.download_button("⬇️ Download Email Log", log_df.to_csv(index=False), file_name="email_log.csv")

    if st.button("🗑️ Clear Email Log"):
        os.remove("email_log.csv")
        st.success("Log cleared.")
else:
    st.info("No logs available yet.")
