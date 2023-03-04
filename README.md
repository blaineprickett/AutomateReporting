# AutomateReporting

This is my workflow for automating reading and writing to a daily excel report in my current role as a food service Retail Manager. (Looking to transition into a developer role this year) Time is all I have. Plus I hate emailing files.

Tools: Microsoft Outlook rules, VBA Scrips, Python, Openpyxl Library, and Microsoft OneDrive.



One time set up -
1)  I set up financial and labor excel reports to be emailed to me overnight. (Kronos Workforce Central & Blackboard) 
2)  Set up Microsoft rules to organize emails in folders on Microsoft Outlook and include VBA Scripts to download the attachments to specified folders in my Microsoft OneDrive account accessible from anywhere. 

Daily -
1)  Set the run time for "ScheduledPythonFiles.py" for the following morning prior to arriving to work (8:30 AM CST) and run the script. Several of the scripts convert each of the downloaded attachments from either .xls to .xlsx or .csv to .xlsx in order to work with the Openpyxl library in Python. The last script displays the newly converted file names for easy copying and pasting.  
2)  Upon arriving to work, I copy and paste each of the newly converted file names into "WriteToReport.py". The names are listed on my screen after scheduling and running "ScheduledPythonFiles.py" overnight.
3)  Set up the cells for which you want to copy from and write to. If the data is in the same cells of the reports each day, you're in great shape. 
4)  Run "WriteToReport.py". All data is copied to the excel report (from 6 different files in my case) and the file is saved to the OneDrive Directory / SharePoint I created for the team. Now my leadership team has access, and I didn't have to send one email. 



VOILA!
