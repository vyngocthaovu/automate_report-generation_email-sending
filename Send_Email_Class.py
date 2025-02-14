'''
The goal of this script is to...
    - Automate the weekly report generation process.
    - Identify and remove duplicate records by comparing the current report with the previous one.
    - Store the report in the designated folder for easy access.
    - Automatically send the report to designated recipients.

This script includes below classes:
1. class RunQuery is to run the SQL query.
2. class CheckDuplicates is to to check the duplicate records and remove if any.
3. class ExportFormat is to export to the excel, and format the file.
4. class SendEmail is to send an email.
5. class Main is to run the entire job.

Author: Vy Vu
Created Date: 10/31/2024
Modified Date: 11/12/2024
'''



import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders



class SendEmail:

    def __init__(self, sender_email, sender_password, recipient_emails, start_str, end_str, have_data=False, path_n_file=None):

        self.sender_email = sender_email
        self.sender_password = sender_password
        self.recipient_emails = recipient_emails

        self.start_str = start_str
        self.end_str = end_str

        self.have_data = have_data
        self.path_n_file = path_n_file


    def send_email(self):

        # Create an email
        msg = MIMEMultipart()
        msg['From'] = self.sender_email
        msg['To'] = ', '.join(self.recipient_emails)
        msg['Subject'] = 'Braille Print Report From ' + self.start_str + ' To ' + self.end_str

        if not self.have_data:
            # Add email content
            message = """
            Hi,

            The report has no records for this period of time.
            ** Do not reply to this email. If you find any inquiries, please contact Vy at staff1@company.org. Thank you.

            Best,
            """

            msg_content = MIMEText(message, _subtype='plain')
            msg.attach(msg_content)

        else:
            # Add email content
            message = """
            Hi,

            Please see attached weekly report.
            ** Do not reply to this email. If you have any inquiries, please contact Vy at staff1@company.org. Thank you.

            Best,
            """

            msg_content = MIMEText(message, _subtype='plain')
            msg.attach(msg_content)
            # Attach report to the email
            part = MIMEBase(_maintype='application', _subtype='octet-stream')
            part.set_payload(open(self.path_n_file, 'rb').read())
            encoders.encode_base64(part)
            part.add_header(_name='Content-Disposition', _value=f"attachment; filename={os.path.basename(self.path_n_file)}")
            msg.attach(part)

        # Send an email
        email_server = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
        email_server.starttls()
        email_server.login(self.sender_email, self.sender_password)
        email_server.sendmail(self.sender_email, self.recipient_emails, msg.as_string())
        email_server.quit()