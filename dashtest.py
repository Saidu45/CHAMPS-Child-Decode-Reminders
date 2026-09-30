"""
CHILD Decode Automated Reminder System
Sends emails 7 days before each activity
No Excel file needed - all dates calculated from calendar logic
"""

import pandas as pd
from datetime import datetime, timedelta
import calendar
import os
import sys

# Try to import email functions
try:
    from outlook_email import send_email
    print("✅ outlook_email imported successfully")
except ImportError as e:
    print(f"❌ Failed to import outlook_email: {e}")
    print("Please make sure outlook_email.py exists")
    # Fallback to print mode
    def send_email(to, subject, body):
        print(f"\n📧 WOULD SEND EMAIL TO: {to}")
        print(f"SUBJECT: {subject}")
        print(f"BODY:\n{body[:200]}...")
        return True

# -----------------------------------
# Core Calendar Functions
# -----------------------------------

def last_saturday(year, month):
    """Return the last Saturday of the given month/year"""
    last_day = calendar.monthrange(year, month)[1]
    date = datetime(year, month, last_day)
    while date.weekday() != 5:  # 5 = Saturday
        date -= timedelta(days=1)
    return date

def get_decode_dates(start_month=None, num_months=6):
    """Get CHILD Decode dates for the next N months"""
    if start_month is None:
        start_month = datetime.now()
    
    dates = []
    year = start_month.year
    month = start_month.month
    
    for i in range(num_months):
        decode_date = last_saturday(year, month)
        dates.append(decode_date)
        
        # Move to next month
        month += 1
        if month > 12:
            month = 1
            year += 1
    
    return dates

def get_next_decode(today=None):
    """Find the next upcoming decode date"""
    today = today or datetime.now().date()
    today = pd.Timestamp(today).normalize()
    
    current_year = today.year
    current_month = today.month
    
    for month_offset in range(0, 13):
        month = current_month + month_offset
        year = current_year
        while month > 12:
            month -= 12
            year += 1
        
        decode_date = last_saturday(year, month)
        if decode_date >= today:
            return decode_date

# -----------------------------------
# Activity Generation
# -----------------------------------

def generate_activities(decode_date):
    """
    Generate all activities for a CHILD Decode cycle
    Based on the given decode date (last Saturday of the month)
    """
    # Calculate activity dates relative to decode date
    # Decode is on a Saturday
    decode_day = decode_date.day
    
    # Monday of decode week (6 days before decode)
    decode_monday = decode_date - timedelta(days=6)
    
    # Thursday of decode week (2 days before decode)
    decode_thursday = decode_date - timedelta(days=2)
    
    # Week after decode
    week_after = decode_date + timedelta(days=7)
    
    # Activities for this decode cycle
    activities = [
        {
            "ID": f"CHILD-{decode_date.strftime('%Y%m')}-001",
            "DATE": decode_date - timedelta(days=30),  # 30 days before decode
            "ASSIGNED PERSON": "Rashid",
            "ACTIVITY": "Rashid sends list to Network Pathologists",
            "DAYS_BEFORE": 30
        },
        {
            "ID": f"CHILD-{decode_date.strftime('%Y%m')}-002",
            "DATE": decode_date - timedelta(days=22),  # 22 days before decode
            "ASSIGNED PERSON": "Drs. Andrew/Aziz",
            "ACTIVITY": "Case packets sent to SMEs",
            "DAYS_BEFORE": 22
        },
        {
            "ID": f"CHILD-{decode_date.strftime('%Y%m')}-003",
            "DATE": decode_date - timedelta(days=12),  # 12 days before decode
            "ASSIGNED PERSON": "Rashid/ Dr. Andrew/Aziz/Seyi",
            "ACTIVITY": "SMEs submit Decode Reports",
            "DAYS_BEFORE": 12
        },
        {
            "ID": f"CHILD-{decode_date.strftime('%Y%m')}-004",
            "DATE": decode_monday,  # Monday of decode week
            "ASSIGNED PERSON": "Drs. Andrew/Aziz",
            "ACTIVITY": "Stakeholder memo sent for Dr. Ike's signature",
            "DAYS_BEFORE": 5
        },
        {
            "ID": f"CHILD-{decode_date.strftime('%Y%m')}-005",
            "DATE": decode_thursday,  # Thursday of decode week
            "ASSIGNED PERSON": "Rashid",
            "ACTIVITY": "Consensus Form sent by Rashid",
            "DAYS_BEFORE": 2
        },
        {
            "ID": f"CHILD-{decode_date.strftime('%Y%m')}-006",
            "DATE": decode_date,  # Decode day (Saturday)
            "ASSIGNED PERSON": "Pathology Team",
            "ACTIVITY": "CHILD Decode",
            "DAYS_BEFORE": 0
        },
        {
            "ID": f"CHILD-{decode_date.strftime('%Y%m')}-007",
            "DATE": decode_date + timedelta(days=2),  # Monday after decode
            "ASSIGNED PERSON": "Drs. Andrew/Aziz/Bassey",
            "ACTIVITY": "Service Completion Certificate sent",
            "DAYS_AFTER": 2
        },
        {
            "ID": f"CHILD-{decode_date.strftime('%Y%m')}-008",
            "DATE": week_after,  # One week after decode
            "ASSIGNED PERSON": "Drs. Andrew/Aziz",
            "ACTIVITY": "Upload Decode Report to REDCap",
            "DAYS_AFTER": 7
        },
        {
            "ID": f"CHILD-{decode_date.strftime('%Y%m')}-009",
            "DATE": week_after,  # One week after decode
            "ASSIGNED PERSON": "Drs. Andrew/Aziz",
            "ACTIVITY": "Send Decode Results to Surveillance",
            "DAYS_AFTER": 7
        }
    ]
    
    return pd.DataFrame(activities)

def generate_all_activities(num_months=6):
    """Generate activities for multiple decode cycles"""
    today = datetime.now()
    decode_dates = get_decode_dates(num_months=num_months)
    
    all_activities = []
    for decode_date in decode_dates:
        # Only include future or very recent activities
        if decode_date + timedelta(days=7) >= today:
            df = generate_activities(decode_date)
            all_activities.append(df)
    
    if all_activities:
        return pd.concat(all_activities, ignore_index=True)
    else:
        return pd.DataFrame()

# -----------------------------------
# Email Directory
# -----------------------------------

EMAILS = {
    "Rashid": "rashid@emory.edu",
    "Drs. Andrew/Aziz": "amosera@emory.edu;saaziz2@emory.edu",
    "Drs. Aziz": "saaziz2@emory.edu",
    "Seyi": "obalog2@emory.edu",
    "Pathology Team": "adaram2@emory.edu",
    "Drs. Bassey": "ibassey@emory.edu",
    "Rashid/ Dr. Andrew/Aziz/Seyi": "rashid@emory.edu;amosera@emory.edu;saaziz2@emory.edu;obalog2@emory.edu",
    "Drs. Andrew/Aziz/Bassey": "amosera@emory.edu;saaziz2@emory.edu;ibassey@emory.edu"
}

def get_email_recipients(assigned_person):
    """Get email addresses for an assigned person or team"""
    return EMAILS.get(assigned_person, None)

# -----------------------------------
# Reminder Engine
# -----------------------------------

def check_and_send_reminders(test_mode=True):
    """
    Check all activities and send reminders 7 days before the activity date
    test_mode: If True, adds [TESTING] to subject and doesn't mark as sent
    """
    
    print("=" * 60)
    print("CHILD Decode Reminder System")
    if test_mode:
        print("⚠️  TEST MODE: Emails will have [TESTING] tag")
    print("=" * 60)
    print()
    
    # Generate activities
    print("📊 Generating activities...")
    df = generate_all_activities(num_months=6)
    
    if df.empty:
        print("❌ No activities generated")
        return
    
    print(f"✅ Generated {len(df)} activities\n")
    
    # Calculate today
    today = pd.Timestamp.today().normalize()
    print(f"📅 Today: {today.strftime('%Y-%m-%d')}")
    
    # Calculate reminder date (7 days before activity)
    reminder_window = 7  # Send reminders 7 days before
    
    # Find activities that need reminders
    df["REMINDER_DATE"] = df["DATE"] - timedelta(days=reminder_window)
    df["DAYS_UNTIL"] = (df["DATE"] - today).dt.days
    df["IS_TODAY"] = df["REMINDER_DATE"].dt.date == today.date()
    
    # Activities that need reminders today
    due_today = df[
        (df["IS_TODAY"]) &
        (df["DATE"] >= today)  # Only future or today's activities
    ].copy()
    
    print(f"📊 Activities due in next {reminder_window} days: {len(due_today)}")
    
    if due_today.empty:
        print("\n✅ No reminders need to be sent today.")
        print("\n📋 Upcoming activities:")
        upcoming = df[df["DATE"] >= today].sort_values("DATE").head(10)
        if not upcoming.empty:
            for _, row in upcoming.iterrows():
                days = (row["DATE"] - today).days
                print(f"  - {row['ACTIVITY']} ({days} days from now)")
        return
    
    print(f"\n📧 Sending {len(due_today)} reminder(s)...")
    print("-" * 60)
    
    emails_sent = 0
    failed_emails = 0
    
    for _, row in due_today.iterrows():
        print(f"\n📧 Activity: {row['ID']}")
        print(f"   Task: {row['ACTIVITY']}")
        print(f"   Activity Date: {row['DATE'].strftime('%d %B %Y')}")
        print(f"   Days until: {row['DAYS_UNTIL']} days")
        print(f"   Assigned to: {row['ASSIGNED PERSON']}")
        
        recipient = get_email_recipients(row["ASSIGNED PERSON"])
        
        if not recipient:
            print(f"   ❌ No email configured for '{row['ASSIGNED PERSON']}'")
            print(f"   📧 Available: {list(EMAILS.keys())}")
            failed_emails += 1
            continue
        
        # Create email
        if test_mode:
            subject = f"[TESTING] CHILD Decode Reminder: {row['ACTIVITY']}"
        else:
            subject = f"CHILD Decode Reminder: {row['ACTIVITY']}"
        
        body = f"""
Dear Team Member,

This is a reminder from the CHAMPS CHILD Decode Management System.

Activity: {row['ACTIVITY']}
Assigned Person: {row['ASSIGNED PERSON']}
Activity Date: {row['DATE'].strftime('%d %B %Y')}
Days Remaining: {row['DAYS_UNTIL']} days

This activity is part of the CHILD Decode cycle for {row['DATE'].strftime('%B %Y')}.

Please ensure this task is completed according to the schedule.

Regards,
CHAMPS Data Management Team
"""
        
        if test_mode:
            body = f"""
⚠️ THIS IS A TEST EMAIL - No action required ⚠️

{body}
"""
        
        try:
            print(f"   📧 Sending to: {recipient}")
            send_email(recipient, subject, body)
            print(f"   ✅ Email sent successfully!")
            emails_sent += 1
            
        except Exception as e:
            print(f"   ❌ Failed: {str(e)}")
            failed_emails += 1
    
    print("\n" + "=" * 60)
    print(f"📊 REMINDER CHECK COMPLETED")
    print(f"✅ Emails sent: {emails_sent}")
    print(f"❌ Failed: {failed_emails}")
    print("=" * 60)
    
    # Show upcoming activities
    print("\n📋 Upcoming activities (next 14 days):")
    upcoming = df[
        (df["DATE"] >= today) &
        (df["DATE"] <= today + timedelta(days=14))
    ].sort_values("DATE")
    
    if not upcoming.empty:
        for _, row in upcoming.iterrows():
            days = (row["DATE"] - today).days
            reminder_date = row["DATE"] - timedelta(days=reminder_window)
            print(f"  - {row['ACTIVITY']}")
            print(f"    Date: {row['DATE'].strftime('%d %B %Y')} ({days} days)")
            print(f"    Reminder: {reminder_date.strftime('%d %B %Y')}")
            print()

# -----------------------------------
# Main
# -----------------------------------

if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("CHILD Decode Reminder System")
    print("=" * 60)
    print()
    
    # Check if we should send real emails
    test_mode = True  # Start in test mode
    
    if test_mode:
        print("🔧 Running in TEST MODE")
        print("   - [TESTING] will be added to subject lines")
        print("   - You can verify emails are being sent correctly")
        print()
        confirm = input("Continue with test? (yes/no): ")
        if confirm.lower() != "yes":
            print("❌ Cancelled")
            sys.exit()
    else:
        print("🔴 Running in PRODUCTION MODE")
        print("   - Real emails will be sent to all recipients")
        print()
        confirm = input("Are you sure? Type 'PRODUCTION' to confirm: ")
        if confirm != "PRODUCTION":
            print("❌ Cancelled")
            sys.exit()
    
    check_and_send_reminders(test_mode=test_mode)