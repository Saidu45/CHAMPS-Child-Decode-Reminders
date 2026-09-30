"""
CHILD Decode Daily Reminder System

- First reminder: 1 day before the activity.
- Sends at most one reminder per day.
- Continues while STATUS is not Completed.
- Stops automatically after the scheduled date.
- TEST_MODE=True adds [TESTING] to subjects.
"""

import os
import pandas as pd
from outlook_email import send_email

FILE = "CHILD_Decode_Tracker.xlsx"
SHEET = "Decode Tasks"

REMINDER_DAYS_BEFORE = 1
TEST_MODE = True

EMAILS = {
    "Rashid": "rashid@emory.edu",
    "Drs. Andrew/Aziz": "amosera@emory.edu;saaziz2@emory.edu",
    "Drs. Aziz": "saaziz2@emory.edu",
    "Seyi": "obalog2@emory.edu",
    "Pathology Team": "adaram2@emory.edu",
    "Drs. Bassey": "ibassey@emory.edu",
    "Rashid / Dr. Andrew / Aziz / Seyi":
        "rashid@emory.edu;amosera@emory.edu;saaziz2@emory.edu;obalog2@emory.edu",
    "Drs. Andrew/Aziz/Bassey":
        "amosera@emory.edu;saaziz2@emory.edu;ibassey@emory.edu",
}

def get_person_name(assigned_person):
    if "/" in assigned_person:
        first_person = assigned_person.split("/")[0].strip()
        if "Drs." in first_person:
            first_person = "Dr. " + first_person.replace("Drs.", "").strip()
        return first_person
    if "Team" in assigned_person:
        return assigned_person
    return assigned_person

def create_email_body(row, today):
    person_name = get_person_name(row["ASSIGNED PERSON"])
    activity_date = row["DATE"].strftime("%d %B %Y")
    days_until = (row["DATE"] - today).days

    if days_until == 0:
        timing = "This activity is due TODAY."
    elif days_until == 1:
        timing = "This activity is due TOMORROW."
    else:
        timing = f"This activity is due in {days_until} day(s)."

    body = f"""Dear {person_name},

This is a reminder from the CHAMPS CHILD Decode Management System.

Activity:
{row['ACTIVITY']}

Assigned Person:
{row['ASSIGNED PERSON']}

Scheduled Date:
{activity_date}

Current Status:
{row['STATUS']}

{timing}

Please complete this activity and update its status to "Completed"
in the CHILD Decode Management System.

Reminders will continue daily until the activity is marked as Completed,
but will automatically stop after the scheduled date.

Best Regards,
CHAMPS Data Management Team
"""

    if TEST_MODE:
        body = "THIS IS A TEST EMAIL - No action required.\n\n" + body

    return body

def send_reminders():
    print("=" * 70)
    print("CHILD Decode Reminder System - DAILY MODE")
    print("=" * 70)

    if not os.path.exists(FILE):
        print(f"File not found: {FILE}")
        return

    try:
        df = pd.read_excel(FILE, sheet_name=SHEET)
    except Exception as exc:
        print(f"Error loading tracker: {exc}")
        return

    df["DATE"] = pd.to_datetime(df["DATE"]).dt.normalize()

    if "STATUS" not in df.columns:
        df["STATUS"] = "Pending"
    if "REMINDER SENT" not in df.columns:
        df["REMINDER SENT"] = "No"
    if "LAST REMINDER SENT" not in df.columns:
        df["LAST REMINDER SENT"] = ""

    df["LAST REMINDER SENT"] = df["LAST REMINDER SENT"].fillna("").astype(str)

    today = pd.Timestamp.today().normalize()
    today_text = today.strftime("%Y-%m-%d")

    reminder_start = df["DATE"] - pd.Timedelta(days=REMINDER_DAYS_BEFORE)

    eligible = df[
        (today >= reminder_start) &
        (today <= df["DATE"]) &
        (df["STATUS"] != "Completed") &
        (df["LAST REMINDER SENT"] != today_text)
    ].copy()

    print(f"Today: {today_text}")
    print(f"Eligible reminders today: {len(eligible)}")

    if eligible.empty:
        print("No reminders need to be sent today.")
        return

    sent = 0
    failed = 0

    for index, row in eligible.iterrows():
        recipient = EMAILS.get(row["ASSIGNED PERSON"], "").strip()

        if not recipient:
            print(f"No email configured for '{row['ASSIGNED PERSON']}'")
            failed += 1
            continue

        subject_prefix = "[TESTING] " if TEST_MODE else ""
        subject = subject_prefix + f"CHILD Decode Reminder: {row['ACTIVITY']}"
        body = create_email_body(row, today)

        try:
            print(f"Sending to: {recipient}")
            send_email(recipient, subject, body)

            df.loc[index, "LAST REMINDER SENT"] = today_text
            df.loc[index, "REMINDER SENT"] = "Yes"
            sent += 1
            print("Reminder sent successfully.")

        except Exception as exc:
            failed += 1
            print(f"Failed to send reminder: {exc}")

    if sent:
        df.to_excel(FILE, sheet_name=SHEET, index=False)

    print("=" * 70)
    print(f"Emails sent: {sent}")
    print(f"Failed: {failed}")
    print("=" * 70)

if __name__ == "__main__":
    send_reminders()
