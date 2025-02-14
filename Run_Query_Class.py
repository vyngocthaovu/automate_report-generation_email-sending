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



import urllib.parse
from sqlalchemy import create_engine
import pandas as pd



class RunQuery:

    def __init__(self, server, database, login, password, driver):
        self.server = server
        self.database = database
        self.login = login
        self.password = password
        self.driver = driver


    def run_query(self):

        # Connect the server and database
        connection_string = f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.login};PWD={self.password}'
        params = urllib.parse.quote_plus(connection_string)
        engine = create_engine(f'mssql+pyodbc:///?odbc_connect={params}')

        # Run the query
        query = """
        SELECT t1.RiderID
            , FullName
            , EffectiveFrom
            , EffectiveThru
            , PrintFormat
        FROM tblCustomer t1
        JOIN tblEligibility t2
        ON t1.CustomerGUID = t2.RiderGUID
        WHERE PrintFormat = 'Braille'
        AND EffectiveFrom >= CAST(DATEADD(WEEK, DATEDIFF(WEEK, 0, GETDATE()) - 2, 0) AS DATE)  -- Sunday two weeks ago
        AND EffectiveFrom < CAST(DATEADD(WEEK, DATEDIFF(WEEK, 0, GETDATE()), 0) AS DATE)  -- Last Sunday (exclusive)
        AND InterviewType NOT IN ('Appeal', 'Extension')
        AND t2.EligibilityStatus NOT IN ('Presumptive', 'Visitor')
        """

        braille_df = pd.read_sql_query(query, engine)

        return braille_df