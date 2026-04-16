# 📧 Cold Email Automation System

Simple, secure Python tool to send personalized cold emails and automated follow-ups using the Gmail API.  
Built for Talent Acquisition — works with CSV contact lists, HTML templates, attachments, and rate limiting.

---

## 🚀 Features

- Personalized intro emails with resume & cover letter attachments  
- Automatic follow-up emails (properly threaded replies)  
- Rate limiting, daily cap, and dry-run mode for safety  
- Full logging of every sent email  
- Works locally and with **Power Automate Desktop**

---

## 📁 Project Structure

```
coldmailer_py_windows/
├── send.py
├── followup.py
├── contacts.csv
├── requirements.txt
├── run_send.ps1
├── run_followup.ps1
├── templates/
├── assets/
├── logs/
├── .env
├── credentials.json
└── token.json
```

---

## ⚙️ Setup Instructions

### 1. Install Dependencies

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

---

### 2. Google Gmail API Setup (One-time)

1. Go to Google Cloud Console  
2. Create a new project  
3. Enable the Gmail API  
4. Create OAuth client ID (Desktop app)  

Rename downloaded file to `credentials.json` and place it in root.

---

### 3. Create `.env` File

```env
CSV_PATH=contacts.csv
DAILY_CAP=50
RATE_PER_SECOND=1
DRY_RUN=0
DRY_RUN_TO=yourtestemail@gmail.com
ATTACH_RESUME=assets/<Resume_FileName>.pdf
ATTACH_COVER=assets/<Cover_Letter_FileName>.pdf
FOLLOWUP_TEMPLATE=followup_v1
```

---

### 4. Prepare Contacts

Edit `contacts.csv` with your leads.

---

## ▶️ How to Run

### Send Emails
```powershell
.\run_send.ps1
```

### Follow-ups
```powershell
.\run_followup.ps1
```

---

## 🤖 Power Automate Desktop

Use "Run PowerShell Script" with your project path and Python executable.

---

## 🔒 Security

- Do NOT commit `token.json`
- Do NOT share `credentials.json`
- Add `.env` and `token.json` to `.gitignore`

---

## 👨‍💻 Author

Parth Tawde
