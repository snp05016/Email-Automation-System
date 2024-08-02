'''Author: Saumya Patel
this is an email generator that generates emails based on html templates provided by the users. It draws data from excel files to generate emails based
on the subject, reciever, due date and time.
'''

import os
import smtplib # EMAIL
from email.message import EmailMessage #EMAIL
from email.utils import formataddr #EMAIL
from pathlib import Path #FOR PATH OF ENV
from dotenv import load_dotenv #LOADING THE ENV
import pandas as pd #DATA READING
from datetime import datetime # RFOR EADING THE TIME FROM A STRING FORMAT
import schedule # SCHEDULING 
import time 

EXCEL_FILE_NAME = 'Clients.xlsx'

PORT = 587 # port number also can change

EMAIL_SERVER = 'smtp.gmail.com' # original email server is gmail but u can change to outlook, icloud or whatever

#check if env in directory
current_dir = Path(__file__).resolve().parent

envars = current_dir / ".env"

print(f"Loading environment variables from: {envars}")

#loading the env var
load_dotenv(envars)

sender_email = os.getenv("email")
password = os.getenv("password")

print(f"Sender Email: {sender_email}")#to check if its exists or not 
print(f"Password: {password}")#check if exists 

def send_email(recipient_email, subject, name, due_date, due_time, topic):
    '''
    Send an email to the recipient with the given subject, name, due date, due time, and topic.
    '''
    from_who = 'name'
    if not sender_email or not password:
        print("Error: Email or password is not set.")
        return

    # read html
    html_template_path = current_dir / "emailtemp.html"
    with open(html_template_path, 'r', encoding='utf-8') as file:
        html_content = file.read()

    # format
    html_content = html_content.format(name=name, due_date=due_date, due_time=due_time, topic=topic)
    msg = EmailMessage()
    msg['From'] = formataddr((f"this is {from_who}", sender_email))
    msg['To'] = recipient_email
    msg['Subject'] = subject

    plain_text_content = f"""\
    Dear {name},

    We are excited to bring you the latest updates and news. Here are some highlights for this month:

    1. Company Achievements
    Our team has reached several milestones this month. We successfully launched new features that have been highly appreciated by our users.

    2. Upcoming Events
    Don't miss out on our upcoming webinar on the latest industry trends. Mark your calendars for {due_date} at {due_time}.

    3. Featured Article
    Check out our featured article on {topic}. It offers great insights and practical advice for professionals in the field.

    Thank you for being a valued subscriber. Stay tuned for more exciting updates!

    Best regards,
    The Team
    """

    msg.set_content(plain_text_content)
    msg.add_alternative(html_content, subtype='html') # formatted text 

    try: # try except block to check if message went or not 
        with smtplib.SMTP(EMAIL_SERVER, PORT) as server:
            server.starttls()
            server.login(sender_email, password)
            server.sendmail(sender_email, recipient_email, msg.as_string())
        print(f"Email sent to {recipient_email} successfully!")
    except Exception as e:
        print(f"Error sending email to {recipient_email}: {e}")

def schedule_emails(due_datetime, recipient_email, subject, name, due_date, due_time, topic):
    '''Schedule the email sending job at the specified due datetime.'''
    def job():
        send_email(
            recipient_email=recipient_email,
            subject=subject,
            name=name,
            due_date=due_date,
            due_time=due_time,
            topic=topic
        )
    
    schedule_time = due_datetime.time().strftime('%H:%M')
    schedule.every().day.at(schedule_time).do(job)# for sending email everyday at the specified due datetime
    schedule.every().once.at(schedule_time).do(job)#optional schedule to send emails only once

def process_excel_file(file_path):
    '''Process the excel file and schedule email sending jobs for each row.'''
    df = pd.read_excel(file_path, dtype={'Due Time': str})#pandas to read excel

    for index, row in df.iterrows():
        due_date_str = row['Due Date']
        due_time_str = row['Due Time']
        due_datetime_str = f"{due_date_str} {due_time_str}"  # due_datetime will formatted as date + time for that date

        try:
            due_datetime = datetime.strptime(due_datetime_str, '%Y-%m-%d %H:%M')
        except ValueError:
            print(f"Date and Time format error for row {index}: {due_datetime_str}")  # error if any for rows
            continue    

        schedule_emails(due_datetime, row['Recipient Email'], row['Subject.'], row['Name'], row['Due Date'], row['Due Time'], row['Topic'])

if __name__ == '__main__':
    process_excel_file(EXCEL_FILE_NAME)

    while True:
        schedule.run_pending()
        time.sleep(1)
