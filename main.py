import pandas as pd
import smtplib
import time
import os
from dotenv import load_dotenv
from email_validator import validate_email, EmailNotValidError
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


# ==============================
# SETTINGS
# ==============================

CONTACTS_FILE = "contacts 2.xlsx"
RESUME_FILE = "Mihir_Kulkarni_Resume.pdf"

DELAY_SECONDS = 8
MAX_EMAILS_TO_SEND = 500  # Change this to 10 for testing, 50/100 for batches


# ==============================
# LOAD ENVIRONMENT VARIABLES
# ==============================

load_dotenv()

EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

if not EMAIL_ADDRESS or not EMAIL_PASSWORD:
    raise ValueError("EMAIL_ADDRESS or EMAIL_PASSWORD is missing in your .env file")


# ==============================
# SMTP CONNECT FUNCTION
# ==============================

def connect_smtp():
    server = smtplib.SMTP("smtp.gmail.com", 587, timeout=60)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    return server


# ==============================
# EMAIL VALIDATION FUNCTION
# ==============================

def is_valid_email_format(email):
    try:
        validate_email(email, check_deliverability=False)
        return True
    except EmailNotValidError:
        return False


# ==============================
# LOAD CONTACTS
# ==============================

df = pd.read_excel(CONTACTS_FILE)

# Clean column names
df.columns = df.columns.astype(str).str.strip().str.lower()

print("Excel columns found:", df.columns.tolist())

required_columns = ["name", "designation", "company", "mail"]

for col in required_columns:
    if col not in df.columns:
        raise ValueError(f"Missing column in Excel file: {col}")

# Remove rows with blank required fields
df = df.dropna(subset=required_columns)

for col in required_columns:
    df[col] = df[col].astype(str).str.strip()

df = df[
    (df["name"] != "") &
    (df["designation"] != "") &
    (df["company"] != "") &
    (df["mail"] != "")
]

print(f"Total usable contacts found: {len(df)}")


# ==============================
# LOAD RESUME
# ==============================

if not os.path.exists(RESUME_FILE):
    raise FileNotFoundError(f"Resume file not found: {RESUME_FILE}")

with open(RESUME_FILE, "rb") as f:
    resume_data = f.read()


# ==============================
# LOG LISTS
# ==============================

sent_logs = []
failed_logs = []
skipped_logs = []

sent_emails = set()


# ==============================
# CONNECT TO SMTP
# ==============================

server = connect_smtp()


# ==============================
# SEND EMAILS
# ==============================

emails_sent_count = 0

for index, row in df.iterrows():

    if emails_sent_count >= MAX_EMAILS_TO_SEND:
        print(f"🛑 Batch limit reached: {MAX_EMAILS_TO_SEND} emails")
        break

    name = row["name"]
    designation = row["designation"]
    company = row["company"]
    email = row["mail"].lower().strip()

    # Skip duplicate emails
    if email in sent_emails:
        print(f"⚠️ Skipping duplicate: {email}")

        skipped_logs.append({
            "name": name,
            "designation": designation,
            "company": company,
            "mail": email,
            "reason": "Duplicate email"
        })

        continue

    sent_emails.add(email)

    # Basic email format validation
    if not is_valid_email_format(email):
        print(f"⚠️ Skipping invalid email format: {email}")

        skipped_logs.append({
            "name": name,
            "designation": designation,
            "company": company,
            "mail": email,
            "reason": "Invalid email format"
        })

        continue

    subject = "Internship Application – Mihir Kulkarni"

    html_body = f"""
    <html>
    <body style="font-family: Arial, sans-serif; color: #333333; line-height: 1.6;">

        <p>Hello {name},</p>

        <p>Hope you are doing well.</p>

        <p>
            I am <strong>Mihir Kulkarni</strong>, a Computer Engineering student pursuing 
            Bachelor of Technology in Computer Science and Engineering.
        </p>

        <p>
            I wish to propose my candidacy for an <strong>internship opportunity</strong> at 
            <strong>{company}</strong>.
        </p>

        <p>
            I enjoy improving systems and applying my learnings to real-world applications. 
            My interests lie in <strong>software development</strong>, 
            <strong>data analytics</strong>, <strong>UI/UX design</strong>, and 
            <strong>business-oriented technology solutions</strong>.
        </p>

        <p>
            I have worked on <strong>Miraki</strong>, a full-stack digital art marketplace 
            built using <strong>React, MongoDB, Node.js</strong>, and deployed via 
            <strong>Vercel</strong>.
        </p>

        <p>
            Apart from technical projects, I have also contributed to event coordination, 
            branding, visual communication, and outreach campaigns. I also served as 
            <strong>Design Co-Head at CSI-RAIT</strong>, where I helped organize UI/UX 
            workshops and managed branding for large-scale tech events.
        </p>

        <p>
            I am reaching out to you as <strong>{designation}</strong> at 
            <strong>{company}</strong>, and I would be grateful if you could consider my 
            profile for any suitable internship opportunity within your team or organization.
        </p>

        <p>
            I have attached my resume for your reference. If you find me a deserving candidate, 
            I would be glad to connect further and discuss how I can contribute to 
            <strong>{company}</strong>.
        </p>

        <p>Hoping to hear from you soon.</p>

        <p style="margin-top: 20px;">
            Warm regards,<br>
            <strong>Mihir Kulkarni</strong><br>
            <a href="mailto:mih.kul.work@gmail.com">mih.kul.work@gmail.com</a><br>
            <a href="https://linkedin.com/in/mihir2005">linkedin.com/in/mihir2005</a><br>
            <a href="https://github.com/Mihirkool">github.com/Mihirkool</a>
        </p>

    </body>
    </html>
    """

    msg = MIMEMultipart()
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = email
    msg["Subject"] = subject

    msg.attach(MIMEText(html_body, "html"))

    resume = MIMEApplication(resume_data, _subtype="pdf")
    resume.add_header(
        "Content-Disposition",
        "attachment",
        filename=RESUME_FILE
    )
    msg.attach(resume)

    try:
        failed_recipients = server.sendmail(EMAIL_ADDRESS, email, msg.as_string())

        if failed_recipients:
            print(f"❌ Failed immediately: {email}")

            failed_logs.append({
                "name": name,
                "designation": designation,
                "company": company,
                "mail": email,
                "reason": str(failed_recipients)
            })

        else:
            print(f"✅ Sent to {name} - {email}")

            sent_logs.append({
                "name": name,
                "designation": designation,
                "company": company,
                "mail": email,
                "status": "Sent"
            })

            emails_sent_count += 1

    except smtplib.SMTPServerDisconnected:
        print(f"⚠️ SMTP disconnected. Reconnecting and retrying: {email}")

        try:
            server = connect_smtp()
            failed_recipients = server.sendmail(EMAIL_ADDRESS, email, msg.as_string())

            if failed_recipients:
                print(f"❌ Failed after reconnect: {email}")

                failed_logs.append({
                    "name": name,
                    "designation": designation,
                    "company": company,
                    "mail": email,
                    "reason": str(failed_recipients)
                })

            else:
                print(f"✅ Sent after reconnect to {name} - {email}")

                sent_logs.append({
                    "name": name,
                    "designation": designation,
                    "company": company,
                    "mail": email,
                    "status": "Sent after reconnect"
                })

                emails_sent_count += 1

        except Exception as e:
            print(f"❌ Failed after reconnect to {email}: {e}")

            failed_logs.append({
                "name": name,
                "designation": designation,
                "company": company,
                "mail": email,
                "reason": str(e)
            })

    except Exception as e:
        print(f"❌ Failed to send to {email}: {e}")

        failed_logs.append({
            "name": name,
            "designation": designation,
            "company": company,
            "mail": email,
            "reason": str(e)
        })

    time.sleep(DELAY_SECONDS)


# ==============================
# CLOSE SMTP SERVER
# ==============================

try:
    server.quit()
except:
    pass


# ==============================
# SAVE LOG FILES
# ==============================

if sent_logs:
    pd.DataFrame(sent_logs).to_excel("sent_emails.xlsx", index=False)

if failed_logs:
    pd.DataFrame(failed_logs).to_excel("failed_emails.xlsx", index=False)

if skipped_logs:
    pd.DataFrame(skipped_logs).to_excel("skipped_emails.xlsx", index=False)

print("✅ Email sending process completed.")
print(f"✅ Sent: {len(sent_logs)}")
print(f"❌ Failed: {len(failed_logs)}")
print(f"⚠️ Skipped: {len(skipped_logs)}")