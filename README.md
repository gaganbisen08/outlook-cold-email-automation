# Outlook Cold Email Automation

Python-based Outlook email automation tool for sending personalized cold emails using Excel contact lists.

# =========================================================
# CONFIGURATION SECTION
# =========================================================

# Excel contact file name
CONTACT_FILE = "demo_hr_contacts.xlsx"

# Resume PDF file name
RESUME_FILE = "demo_resume.pdf"

# Progress tracking file
PROGRESS_FILE = "progress.txt"

# Daily sending limit
DAILY_LIMIT = 75

# Minimum delay between emails (seconds)
MIN_DELAY = 60

# Maximum delay between emails (seconds)
MAX_DELAY = 180

# Enable resume support
ENABLE_RESUME_PROGRESS = True

# Test mode
# True  = sends only to your test email
# False = sends to actual HR emails
TEST_MODE = False

# Test email address
TEST_EMAIL = "yourtestemail@example.com"

# Email signature details
YOUR_NAME = "Demo Candidate"

YOUR_PHONE = "+91 9876543210"

YOUR_EMAIL = "demo.candidate@example.com"

LINKEDIN_URL = "linkedin.com/in/democandidate"

GITHUB_URL = "github.com/democandidate"
## Features

- Outlook desktop integration
- Personalized email generation
- Resume attachment support
- Randomized subject lines
- Human-like sending delays
- Resume progress after interruption
- Excel-based contact management

## Technologies Used

- Python
- pandas
- pywin32
- openpyxl
- Microsoft Outlook

## Setup

### Install Dependencies

```bash
pip install -r requirements.txt
```

### Configure Files

Place:
- contact Excel file
- resume PDF
- Python script

in same folder.

### Run

```bash
python send_mails.py
```

## Use Cases

- HR outreach
- Job applications
- Recruiter cold emailing
- Email workflow automation

## Disclaimer

Use responsibly and follow email sending limits to avoid spam detection.
