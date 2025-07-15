import pandas as pd
import smtplib
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from dotenv import load_dotenv
import os

# Load environment variables
load_dotenv()
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# Load Excel contacts
df = pd.read_excel('contacts.xlsx')

# Set up Gmail server
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

# Load resume once
with open('Mihir_Kulkarni_Resume.pdf', 'rb') as f:
    resume_data = f.read()

# Send one-by-one
for index, row in df.iterrows():
    name = row['name']
    email = row['email']
    company = row['company']
    designation = row['designation'] if 'designation' in row and not pd.isna(row['designation']) else 'your team'

    subject = "Internship Application – Mihir Kulkarni | Data & Design Driven"

    html_body = f"""
    <html>
    <body style="font-family: Arial, sans-serif; color: #333;">
        <p>Dear {name} Sir/Ma'am,</p>

        <p>
            I’m <strong>Mihir Kulkarni</strong>, a Computer Engineering student with a passion for solving real-world problems
            through <strong>data analytics</strong>, <strong>software development</strong>, and <strong>creative design thinking</strong>.
            I'm reaching out to express interest in internship opportunities under your leadership as <strong>{designation}</strong> at <strong>{company}</strong>.
        </p>

        <p>
            Apart from tech, I’ve actively contributed to the business and management side of things. At the recent <strong>Technext</strong> event,
            I played a key role in coordinating lab visits, guiding experts across departments, and ensuring smooth flow of presentations.
            I also worked on growth strategy and visual communication for multiple outreach campaigns — experiences that helped me build strong organizational and business development instincts.
        </p>

        <p>
            On the creative front, I served as <strong>Design Co-Head</strong> at <strong>CSI-RAIT</strong>, organizing UI/UX workshops and managing branding for large-scale tech events.
            I’ve also developed <strong>Miraki</strong> – a full-stack digital art marketplace using <em>React, MongoDB, Node.js</em> and deployed it via <em>Vercel</em>.
        </p>

        <p>
            My experience across both tech and business-focused roles allows me to contribute value in multidisciplinary teams.
            I’ve attached my resume and would be thrilled to connect further if there’s a potential fit.
        </p>

        <p>Thank you for your time and consideration.</p>

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
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = email
    msg['Subject'] = subject
    msg.attach(MIMEText(html_body, 'html'))

    # Attach resume
    resume = MIMEApplication(resume_data, _subtype='pdf')
    resume.add_header('Content-Disposition', 'attachment', filename='Mihir_Kulkarni_Resume.pdf')
    msg.attach(resume)

    try:
        server.sendmail(EMAIL_ADDRESS, email, msg.as_string())
        print(f"✅ Sent to {name} ({email})")
    except Exception as e:
        print(f"❌ Failed to send to {email}: {e}")

    time.sleep(3)

server.quit()
