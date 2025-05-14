# 📧 Email Tone Analyzer – "FBI: Feelings & Behavior Investigator"

This Python-based tool automatically fetches recent emails, analyzes their emotional tone using a transformer-based NLP model (`roberta-base-go_emotions`), and generates an Excel report summarizing detected emotions per message. The tool also supports automated email delivery of the report.

---

## 🚀 Features

- Connects to an IMAP server to fetch recent emails.
- Uses HuggingFace Transformers to detect emotions in email content.
- Cleans and sanitizes email body for accurate NLP processing.
- Detects **primary and secondary emotions** in emails.
- Generates a styled Excel report with analysis.
- Automatically sends the Excel report via email.
- Designed for support teams, HR, or sentiment-driven analytics.

---

## 🧠 Model Used

- `SamLowe/roberta-base-go_emotions`
  - Fine-tuned for emotion classification (joy, anger, fear, etc.)
  - Supports multi-label output (top_k=None)

---

## 📦 Dependencies

Install the required libraries using:

```bash
pip install transformers openpyxl tqdm beautifulsoup4
This project also requires:

imaplib, email, smtplib (Python stdlib)

A valid IMAP email server

SMTP access for sending the report

⚙️ Configuration
Edit the script or set these variables in your environment:

python
Copy
Edit
IMAP_SERVER = "your.imap.server"
EMAIL_ACCOUNT = "your@email.com"
EMAIL_PASSWORD = "yourpassword"
MAILBOX = "INBOX"  # or a nested mailbox like "INBOX/Mails/Support"
🧪 Usage
bash
Copy
Edit
python email_analyzer_summary.py
Or directly in code:

python
Copy
Edit
analyzer = EmailToneAnalyzer(
    imap_server=IMAP_SERVER,
    email_account=EMAIL_ACCOUNT,
    email_password=EMAIL_PASSWORD,
    mailbox=MAILBOX
)
analyzer.run_analysis(days_back=1, recipient_email="manager@example.com")
📊 Output
Excel file: email_tone_YYYY-MM-DD.xlsx

Columns include:

Sender

Subject

Date

Primary Emotion

Score

Secondary Emotions

📤 Email Report
The tool optionally sends the generated report via email with a summary like:

yaml
Copy
Edit
Summary:

- Total emails analyzed: 14

- JOY: 6 emails (43%)
- ANGER: 3 emails (21%)
...
🛡️ Security
Avoid hardcoding credentials. Use environment variables or secure vaults in production.

📅 Automation Tip
Add this script to a daily cron job to automate email tone reporting.

👨‍💼 Sample Signature
🕶️ FBI (Feelings & Behavior Investigator)
We read between the lines — and also judge them.
