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
import xlsxwriter
import pandas as pd



class ExportFormat:

    def __init__(self, df, start_str, end_str, new_report_path):
        
        self.df = df
        
        self.start_str = start_str
        self.end_str = end_str

        self.new_report_path = new_report_path

        self.report_name = self.start_str + '-' + self.end_str + ' BraillePrint.xlsx'
        self.path_n_file = os.path.join(self.new_report_path, self.report_name)
        

    def clean_and_export(self):

        file_writer = pd.ExcelWriter(self.path_n_file)
        self.df.to_excel(file_writer, sheet_name='BraillePrint', index=False)
        workbook = file_writer.book
        worksheet = file_writer.sheets['BraillePrint']

        # Remove the border
        border_fmt = workbook.add_format({'bottom': None, 'top': None, 'left': None, 'right': None})
        worksheet.conditional_format(xlsxwriter.utility.xl_range(0, 0, len(self.df), len(self.df)),
                                        {'type': 'no_errors', 'format': border_fmt})

        # Adjust the font and size of letters, and the width of the column
        font_fmt = workbook.add_format({'font_name': 'Calibri', 'font_size': 10})
        for idx, col in enumerate(self.df):
            col_len = max(self.df[col].astype(str).map(len).max(), len(col))
            worksheet.set_column(first_col=idx, last_col=idx, width=col_len, cell_format=font_fmt)

        file_writer.close()

        return self.path_n_file