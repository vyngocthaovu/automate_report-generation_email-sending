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
import openpyxl
from datetime import datetime, timedelta
from Run_Query_Class import RunQuery
from Check_Duplicates_Class import CheckDuplicates
from Export_Format_Class import ExportFormat
from Send_Email_Class import SendEmail



class Main:

    def __init__(self):

        # Server
        self.server = 'server_name'
        self.database = 'database_name'
        self.login = 'username'
        self.password = 'password'
        self.driver = '{ODBC Driver 17 for SQL Server}'

        # Path
        self.path = '.\\Outputs'

        # Start date and end date of the new report (for report's name only)
        self.today = datetime.today()
        self.start_date = self.today - timedelta(weeks=2, days=self.today.weekday() + 1)  # Sunday two weeks ago
        self.end_date = self.today - timedelta(days=self.today.weekday() + 2)  # Last Saturday
        self.start_str = self.start_date.strftime('%m.%d.%Y')  # Format the dates as string in the format 'MM.DD.YYYY'
        self.end_str = self.end_date.strftime('%m.%d.%Y')

        # Start date of the previous report (to check the duplicate records)
        self.previous_start_date = self.today - timedelta(weeks=3, days=self.today.weekday() + 1)
        self.previous_report_year = self.previous_start_date.year

        # Folder of the previous report and the new report
        ## The previous report
        self.previous_report_path = os.path.join(self.path, str(self.previous_report_year))
        ## The current report
        self.folder_name = str(self.start_date.year)
        self.new_report_path = os.path.join(self.path, self.folder_name)
        if not os.path.exists(self.new_report_path):
            os.makedirs(self.new_report_path)

        # Emails
        self.sender_email = 'sender_email@company.org'
        self.sender_password = 'password'
        self.recipient_emails = ['recipient1@company.org', 'recipient2@company.org']

    
    def run(self):

        # Run the query
        run_query_instance = RunQuery(self.server, self.database, self.login, self.password, self.driver)
        report_df = run_query_instance.run_query()

        if len(report_df) <= 0:  # No data in the new report
            have_data = False
            send_email_instance = SendEmail(self.sender_email, self.sender_password, self.recipient_emails, self.start_str, self.end_str, have_data=have_data)
            send_email_instance.send_email()
        
        else:  # There is data in the new report
            check_duplicates_instance = CheckDuplicates(df=report_df, previous_start_date=self.previous_start_date, previous_report_path=self.previous_report_path)
            have_data, revised_df = check_duplicates_instance.check_n_remove_duplicates()
            
            if not have_data:  # After removing duplicates, the report has no data
                send_email_instance = SendEmail(self.sender_email, self.sender_password, self.recipient_emails, self.start_str, self.end_str, have_data=have_data)
                send_email_instance.send_email()
            else:  # After removing duplicates, the report has data
                export_format_instance = ExportFormat(df=revised_df, start_str=self.start_str, end_str=self.end_str, new_report_path=self.new_report_path)
                path_n_file = export_format_instance.clean_and_export()
                send_email_instance = SendEmail(self.sender_email, self.sender_password, self.recipient_emails, self.start_str, self.end_str, have_data=have_data, path_n_file=path_n_file)
                send_email_instance.send_email()


if __name__ == "__main__":
    main_program = Main()
    main_program.run()