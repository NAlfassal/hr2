import pandas as pd
import sys
import os
import smtplib
from datetime import date
from email.mime.text import MIMEText
from email.header import Header

def check_and_remind(responses_path, master_list_excel_path, gmail_user, gmail_pass, form_url):
    try:
        #  تأكيد عدم الإرسال المكرر في نفس اليوم
        LOCK_FILE = f"reminder_lock_{date.today()}.txt"
        if os.path.exists(LOCK_FILE):
            print("Reminder already sent today. Skipping duplicate.")
            sys.exit(0)
        with open(LOCK_FILE, "w") as f:
            f.write("sent")

        # 1. قراءة قائمة كل الموظفين
        if not os.path.exists(master_list_excel_path):
            print(f"Error: Master employee list not found at {master_list_excel_path}", file=sys.stderr)
            sys.exit(1)
        master_df = pd.read_excel(master_list_excel_path)
        all_employees = set(master_df['Email'].dropna().unique())
        
        # 2. قراءة الردود
        if not os.path.exists(responses_path):
            responded_employees = set()
        else:
            responses_df = pd.read_excel(responses_path)
            if 'البريد الإلكتروني' in responses_df.columns:
                responded_employees = set(responses_df['البريد الإلكتروني'].dropna().unique())
            else:
                responded_employees = set()

        # 3. من لم يرد
        non_responders = all_employees - responded_employees

        if not non_responders:
            print("Success: All employees have responded.")
            return

        print(f"Found {len(non_responders)} employees who have not responded. Sending reminders...")

        # 4. إرسال التذكيرات
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(gmail_user, gmail_pass)

        for email in non_responders:
            subject = "تذكير: الرجاء تعبئة نموذج حصر الأنشطة المعرفية"
            body = f"""مرحبًا،

هذا تذكير لطيف بضرورة تعبئة نموذج حصر الأنشطة المعرفية. الموعد النهائي قريب.

الرجاء استخدام الرابط التالي لتعبئة النموذج:
{form_url}

شكرًا لتعاونكم.
"""
            msg = MIMEText(body, 'plain', 'utf-8')
            msg['Subject'] = Header(subject, 'utf-8')
            msg['From'] = gmail_user
            msg['To'] = email

            server.sendmail(gmail_user, email, msg.as_string())
            print(f"Reminder sent to: {email}")

        server.quit()
        print("All reminders sent successfully.")

    except Exception as e:
        print(f"An error occurred: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) != 6:
        print("Usage: python reminder_checker.py <responses_path> <master_list_excel_path> <gmail_user> <gmail_pass> <form_url>", file=sys.stderr)
        sys.exit(1)
    
    check_and_remind(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5])
