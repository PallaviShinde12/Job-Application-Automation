import os
import re
import base64
import pymupdf

from datetime import datetime
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from openpyxl import load_workbook


# ---------------- SETTINGS ----------------

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

LABEL_ID = "Label_2232474142574291401"

EXCEL_FILE = "Job_Application_Tracker.xlsx"

DOWNLOAD_FOLDER = "."


# ---------------- EXTRACT JOB POSITION FROM SUBJECT ----------------

def extract_position_from_subject(subject):

    subject = subject.strip()

    # Remove "Subject:" from beginning
    subject = re.sub(r'(?i)^subject:\s*', '', subject).strip()

    patterns = [
        r'(?i)application\s+for\s+(.+?)\s+position(?:\s*[-:–|]\s*resume.*)?$',
        r'(?i)resume\s+for\s+(.+?)\s+position$',
        r'(?i)resume\s+for\s+(.+)',
        r'(?i)application\s+for\s+(.+)',
        r'(?i)application\s+as\s+(.+)',
        r'(?i)job\s+application\s+for\s+(.+)',
        r'(?i)job\s+application\s*[-:–]?\s*(.+)',
        r'(?i)applying\s+for\s+(.+)',
        r'(?i)applying\s+as\s+(.+)',
    ]

    for pattern in patterns:

        match = re.search(pattern, subject)

        if match:

            position = match.group(1).strip()

            position = re.sub(
                r'(?i)\s+position$',
                '',
                position
            ).strip()

            position = re.sub(
                r'(?i)\s*[-:–|]\s*(resume|cv).*?$',
                '',
                position
            ).strip()

            return position

    return ""


# ---------------- GMAIL CONNECTION ----------------

creds = Credentials.from_authorized_user_file(
    "token.json",
    SCOPES
)

service = build(
    "gmail",
    "v1",
    credentials=creds
)

print("Gmail connected successfully!")


# ---------------- FIND JOB APPLICATION EMAILS ----------------

results = service.users().messages().list(
    userId="me",
    q="{subject:Application subject:Job subject:Interview subject:Recruitment subject:Hiring subject:Resume} has:attachment"
).execute()

messages = results.get("messages", [])

print("Job Application emails found:", len(messages))


# ---------------- EXCEL ----------------

if not os.path.exists(EXCEL_FILE):

    print("Excel file not found!")
    exit()

workbook = load_workbook(EXCEL_FILE)

sheet = workbook["Job_Applications"]


# ---------------- PROCESS EMAILS ----------------

for message in messages:

    message_id = message["id"]

    print("\nProcessing email:", message_id)

    email_data = service.users().messages().get(
        userId="me",
        id=message_id
    ).execute()


    # ---------------- GET EMAIL SUBJECT ----------------

    headers = email_data.get("payload", {}).get("headers", [])

    subject = next(
        (
            h["value"]
            for h in headers
            if h["name"].lower() == "subject"
        ),
        ""
    )

    applied_for = extract_position_from_subject(subject)

    print("Email Subject:", subject)
    print("Applied For:", applied_for)


    # ---------------- FIND PDF ATTACHMENT ----------------

    payload = email_data.get("payload", {})

    parts = payload.get("parts", [])

    pdf_file = None


    for part in parts:

        filename = part.get("filename", "")

        if filename.lower().endswith(".pdf"):

            attachment_id = part.get("body", {}).get("attachmentId")

            if attachment_id:

                attachment = service.users().messages().attachments().get(
                    userId="me",
                    messageId=message_id,
                    id=attachment_id
                ).execute()

                file_data = base64.urlsafe_b64decode(
                    attachment["data"]
                )

                pdf_file = os.path.join(
                    DOWNLOAD_FOLDER,
                    filename
                )

                with open(pdf_file, "wb") as f:

                    f.write(file_data)

                print("Resume downloaded:", filename)

                break


    if not pdf_file:

        print("No PDF resume found.")

        continue


    # ---------------- READ PDF ----------------

    document = pymupdf.open(pdf_file)

    text = ""

    for page in document:

        text += page.get_text()

    document.close()


    # ---------------- EXTRACT EMAIL ----------------

    email_match = re.search(
        r'[\w\.-]+@[\w\.-]+\.\w+',
        text
    )

    email = email_match.group(0) if email_match else ""


    # ---------------- EXTRACT PHONE ----------------

    phone_match = re.search(
        r'(?:\+91[\s-]?)?[6-9]\d{9}',
        text
    )

    phone = phone_match.group(0) if phone_match else ""


    # ---------------- EXTRACT NAME ----------------

    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    name = ""


    for line in lines:

        if (
            "email" not in line.lower()
            and "phone" not in line.lower()
            and "linkedin" not in line.lower()
            and len(line.split()) <= 4
        ):

            if re.match(r'^[A-Za-z ]+$', line):

                name = line

                break


    # ---------------- SKILLS ----------------

    text_lower = text.lower()

    skills_list = [
        "Excel",
        "Power Query",
        "Power BI",
        "SQL",
        "DAX",
        "MySQL",
        "Python",
        "Java",
        "C",
        "Data Analysis",
        "Data Visualization",
        "Tableau"
    ]

    found_skills = []


    for skill in skills_list:

        if skill.lower() in text_lower:

            found_skills.append(skill)


    skills = ", ".join(found_skills)


    # ---------------- EDUCATION ----------------

    education = ""


    if "b.tech" in text_lower or "b tech" in text_lower:

        education = "B.Tech"

    elif "bachelor" in text_lower:

        education = "Bachelor's Degree"

    elif "m.tech" in text_lower or "m tech" in text_lower:

        education = "M.Tech"

    elif "master" in text_lower:

        education = "Master's Degree"


    # ---------------- EXPERIENCE ----------------

    if "fresher" in text_lower:

        experience = "Fresher"

    elif "internship" in text_lower:

        experience = "Internship"

    elif "experience" in text_lower:

        experience = "Experienced"

    else:

        experience = "Fresher"


    # ---------------- CHECK DUPLICATE ----------------

    duplicate = False


    for row in sheet.iter_rows(
        min_row=2,
        values_only=True
    ):

        existing_email = row[2]

        existing_resume = row[8]


        if email and existing_email == email:

            duplicate = True

            break


        if existing_resume == os.path.basename(pdf_file):

            duplicate = True

            break


    if duplicate:

        print("Candidate already exists. Skipping.")

        continue


    # ---------------- ADD TO EXCEL ----------------

    new_row = [

        datetime.now().date(),

        name,

        email,

        phone,

        skills,

        education,

        experience,

        applied_for,

        os.path.basename(pdf_file),

        "New"
    ]


    sheet.append(new_row)


    print("\n--- CANDIDATE ADDED ---")

    print("Name:", name)

    print("Email:", email)

    print("Phone:", phone)

    print("Applied For:", applied_for)

    print("Skills:", skills)

    print("Education:", education)

    print("Experience:", experience)


# ---------------- SAVE EXCEL ----------------

workbook.save(EXCEL_FILE)


print("\n================================")

print("EXCEL UPDATED SUCCESSFULLY")

print("================================")
