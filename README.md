Yes, this one is much cleaner and shorter. Copy-paste this full code into `README.md` ✅

````markdown
# Mail Automation System

A Python-based bulk email automation project that reads contacts from an Excel file and sends personalized emails using Gmail SMTP.

---

## Features

- Reads contact data from Excel
- Sends personalized emails
- Uses Gmail SMTP
- Attaches resume PDF
- Skips duplicate emails
- Creates sent, failed, and skipped email logs
- Handles SMTP disconnect and reconnect

---

## Project Structure

```text
mail_automation_system/
│
├── main.py
├── Sample.xlsx
├── Mihir_Kulkarni_Resume.pdf
├── requirements.txt
├── .env
├── .gitignore
└── README.md
````

---

## Requirements

Install required libraries:

```bash
pip install -r requirements.txt
```

`requirements.txt` should contain:

```text
pandas
openpyxl
python-dotenv
```

---

## Excel Format

The Excel file must have exactly these 4 columns:

```text
name | designation | company | mail
```

Example from `Sample.xlsx`:

| name     | designation    | company | mail                                    |
| -------- | -------------- | ------- | --------------------------------------- |
| person 1 | Sr. HR manager | a       | [xyz@gmail.com](mailto:xyz@gmail.com)   |
| person 2 | Tech HR        | b       | [abc@gmail.com](mailto:abc@gmail.com)   |
| person 3 | Jr. HR intern  | c       | [test@gmail.com](mailto:test@gmail.com) |

In `main.py`, the file is selected using:

```python
CONTACTS_FILE = "Sample.xlsx"
```

---

## Resume File

Keep your resume PDF in the project folder.

Default file name:

```text
Mihir_Kulkarni_Resume.pdf
```

In `main.py`, it is selected using:

```python
RESUME_FILE = "Mihir_Kulkarni_Resume.pdf"
```

---

## Gmail App Password Setup

This project uses Gmail SMTP, so you need a Gmail App Password.

Steps:

1. Open your Google Account.
2. Go to **Security**.
3. Enable **2-Step Verification**.
4. Search for **App Passwords**.
5. Create an app password for Mail or custom app.
6. Copy the generated 16-character password.

Do not use your normal Gmail password.

---

## Environment File

Create a `.env` file in the project folder:

```env
EMAIL_ADDRESS=yourgmail@gmail.com
EMAIL_PASSWORD=your_gmail_app_password
```

Example:

```env
EMAIL_ADDRESS=sample@gmail.com
EMAIL_PASSWORD=abcd efgh ijkl mnop
```

---

## Gmail SMTP Details

The script uses Gmail SMTP:

| Field       | Value                   |
| ----------- | ----------------------- |
| SMTP Server | smtp.gmail.com          |
| Port        | 587                     |
| Security    | TLS                     |
| Login       | Gmail ID + App Password |

---

## How to Run

Step-by-step:

```bash
git clone https://github.com/Mihirkool/mail_automation_system.git
cd mail_automation_system
pip install -r requirements.txt
python main.py
```

Before running, make sure these files exist:

```text
main.py
Sample.xlsx
Mihir_Kulkarni_Resume.pdf
.env
```

---

## Sample Email Output

For this Excel row:

| name     | designation    | company | mail                                  |
| -------- | -------------- | ------- | ------------------------------------- |
| person 1 | Sr. HR manager | a       | [xyz@gmail.com](mailto:xyz@gmail.com) |

The email will be personalized like this:

```text
Subject: Internship Application – Mihir Kulkarni

Hello person 1,

Hope you are doing well.

I am Mihir Kulkarni, a Computer Engineering student pursuing Bachelor of Technology in Computer Science and Engineering.

I wish to propose my candidacy for an internship opportunity at a.

I am reaching out to you as Sr. HR manager at a, and I would be grateful if you could consider my profile for any suitable internship opportunity.

I have attached my resume for your reference.

Warm regards,
Mihir Kulkarni
mih.kul.work@gmail.com
linkedin.com/in/mihir2005
github.com/Mihirkool
```

---

## Output Files

After running, the script may create:

```text
sent_emails.xlsx
failed_emails.xlsx
skipped_emails.xlsx
```

* `sent_emails.xlsx` contains successfully sent emails.
* `failed_emails.xlsx` contains failed email attempts.
* `skipped_emails.xlsx` contains duplicate or invalid emails.

---

## Important Code Settings

Inside `main.py`:

```python
DELAY_SECONDS = 10
MAX_EMAILS_TO_SEND = 50
```

For testing:

```python
MAX_EMAILS_TO_SEND = 5
```

For safer bulk sending:

```python
MAX_EMAILS_TO_SEND = 20
DELAY_SECONDS = 15
```

---

## Note About Invalid Emails

The script can detect bad email formats like:

```text
wrongemail.com
abc@
test@gmail
```

But it cannot always detect whether a valid-looking email inbox actually exists.

Example:

```text
fresher.hiring@hcltech.com
```

The format and domain may be valid, but the actual inbox may not exist. In that case, Gmail may send a bounce message later.

---

## Recommended `.gitignore`

```gitignore
.env
*.xlsx
__pycache__/
*.pyc
```

Since Excel files are ignored, add only the safe sample file using:

```bash
git add -f Sample.xlsx
```

Do not upload real contacts or `.env`.

---

## GitHub Update Commands

Use these commands when updating the project:

```bash
git status
git add main.py README.md
git add -f Sample.xlsx
git commit -m "Update project files"
git push origin main
```

---

## Author

Mihir Kulkarni

* Email: [mih.kul.work@gmail.com](mailto:mih.kul.work@gmail.com)
* LinkedIn: linkedin.com/in/mihir2005
* GitHub: github.com/Mihirkool

---

After saving it, push README:

```bash
git add README.md
git commit -m "Update README"
git push origin main
```
