import pandas as pd
import win32com.client as win32
import time
import random
from pathlib import Path

# =========================================================
# LOAD EXCEL
# =========================================================

df = pd.read_excel("xl_sheet name")

# =========================================================
# RESUME PATH
# =========================================================

resume_path = str(Path("your_resume_name").resolve())

# =========================================================
# LOAD LAST POSITION
# =========================================================

progress_file = "progress.txt"

try:
    with open(progress_file, "r") as f:
        start_index = int(f.read().strip())
except:
    start_index = 0

print(f"Resuming from row: {start_index}")

# =========================================================
# START OUTLOOK
# =========================================================

outlook = win32.Dispatch('outlook.application')

# =========================================================
# DAILY LIMIT
# =========================================================

daily_limit = 75
count = 0

# =========================================================
# ENTRY-LEVEL SUBJECT VARIATIONS
# =========================================================

subjects = [

    "replace your subject according to your role",

    "MCA Graduate Applying for Technical Support Role",

    "Application for Fresher IT Support Opportunity",

    "Entry-Level Technical Support Application – Gagan Bisen",

    "Application for Desktop Support / IT Support Role",

    "MCA Graduate | IT Support & Troubleshooting",

    "Application for Service Desk / Technical Support Role",

    "Fresher Application for IT Support Position",

    "Entry-Level IT Support Engineer Application",

    "Technical Support Fresher | MCA Graduate"

]

# =========================================================
# ENTRY-LEVEL INTRO VARIATIONS
# =========================================================

intro_lines = [

    "I recently completed my MCA and am actively seeking an entry-level opportunity in IT Support or Technical Support.",

    "I’m an MCA graduate with practical exposure to IT support, troubleshooting, networking fundamentals, and Windows/Linux environments through hands-on projects and self-learning.",

    "I have built practical projects involving help desk systems, virtualization, automation, and system troubleshooting to strengthen my IT support skills.",

    "My background includes hands-on learning in Windows administration, technical support, networking fundamentals, and IT troubleshooting through lab-based projects.",

    "I recently completed my MCA and have focused on building practical IT support skills involving osTicket, PowerShell automation, and virtualization environments.",

    "I’m looking for an entry-level IT Support or Service Desk opportunity where I can apply my troubleshooting, support, and system administration skills.",

    "I have practical exposure to technical support operations through projects involving Windows/Linux systems, automation scripting, and help desk workflows.",

    "I’ve developed practical knowledge in IT support, networking, remote troubleshooting, and ticketing systems through self-directed projects and technical learning.",

    "I’m an MCA graduate with a strong interest in IT Support, Technical Support, and system troubleshooting, backed by hands-on project experience.",

    "My technical background includes practical projects in help desk operations, virtualization, Windows support, and troubleshooting workflows."

]
# =========================================================
# MAIN LOOP
# =========================================================

for index in range(start_index, len(df)):

    if count >= daily_limit:
        print("Daily limit reached.")
        break

    try:

        row = df.iloc[index]

        name = str(row['Name']).strip()
        email = str(row['Email']).strip()

        if "@" not in email:
            continue

        # =========================================================
        # RANDOM CONTENT
        # =========================================================

        subject = random.choice(subjects)

        intro = random.choice(intro_lines)

        # =========================================================
        # CREATE EMAIL
        # =========================================================

        mail = outlook.CreateItem(0)

        mail.To = email

        mail.Subject = subject

        mail.HTMLBody = f"""
        <p>Dear {name},</p>

        <p>
        {intro}
        </p>

        <p>
        Projects:
        </p>

        <ul>
            <li>osTicket Help Desk System</li>
            <li>PowerShell & Bash Automation Toolkit</li>
            <li>Multi-OS Virtualization Lab</li>
        </ul>

        <p>
        Skilled in:
        Windows Administration, Active Directory,
        Microsoft 365 Support, TCP/IP, DNS, DHCP,
        Remote Desktop Support, and IT Troubleshooting.
        </p>

        <p>
        I’ve attached my resume for your consideration and would appreciate
        the opportunity to discuss relevant openings.
        </p>

        <p>
        Best regards,<br><br>

        <b>Gagan Bisen</b><br>

        📞 +91 1234567890<br>

        📧 demo2@gmail.com<br><br>

        🔗 LinkedIn:<br>
        linkedin.com/in/gagan-bisen<br><br>

        🔗 GitHub:<br>
        github.com/gaganbisen08
        </p>
        """

        # =========================================================
        # ATTACH RESUME
        # =========================================================

        mail.Attachments.Add(resume_path)

        # =========================================================
        # SEND EMAIL
        # =========================================================

        mail.Send()

        count += 1

        print(f"[{count}] Sent to: {email}")

        # =========================================================
        # SAVE PROGRESS
        # =========================================================

        with open(progress_file, "w") as f:
            f.write(str(index + 1))

        # =========================================================
        # RANDOM DELAY
        # =========================================================

        sleep_time = random.randint(60, 180)

        print(f"Waiting {sleep_time} seconds...\n")

        time.sleep(sleep_time)

    except Exception as e:

        print(f"Failed for {email}")
        print(e)

print("Email process completed.")