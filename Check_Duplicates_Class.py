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
import pandas as pd



class CheckDuplicates:

    def __init__(self, df, previous_start_date, previous_report_path):

        self.df = df
        
        self.previous_start_date = previous_start_date
        self.previous_start_str = previous_start_date.strftime('%m.%d.%Y')
        self.previous_report_path = previous_report_path

        self.previous_report_df = None
        self.duplicates = None
        self.duplicateIDs = []

    def check_n_remove_duplicates(self):

        for file in os.listdir(self.previous_report_path):
            if file.startswith(self.previous_start_str):
                self.previous_report_df = pd.read_excel(os.path.join(self.previous_report_path, file), header=0)  # Convert the previous report to DataFrame
                self.previous_report_df['RiderID'] = self.previous_report_df['RiderID'].astype(str)
                self.df['RiderID'] = self.df['RiderID'].astype(str)

                self.duplicates = pd.merge(self.df, self.previous_report_df, on='RiderID')

                if len(self.duplicates) > 0:
                    self.duplicateIDs = self.duplicates.iloc[:, 0].tolist()
                    self.df = self.df[~self.df.iloc[:, 0].isin(self.duplicateIDs)]
                break

        if len(self.df) > 0:
            have_data = True
            return have_data, self.df
        else:
            have_data = False
            return have_data, None