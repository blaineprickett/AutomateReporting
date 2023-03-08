# AutomateReporting

This is my workflow for automating reading and writing to a daily excel report in my current role as a food service Retail Manager. (Looking to transition into a developer role this year) Time is all I have. Plus I hate emailing files.
______________________________________________________
Tools: Microsoft Outlook rules, VBA Scrips, Python, Openpyxl Library, and Microsoft OneDrive.
______________________________________________________
One time set up -
1)  I set up financial and labor excel reports to be emailed to me overnight. (Kronos Workforce Central & Blackboard) These files are delivered as .xls and .csv file types.
2)  Set up Microsoft Outlook rules to organize emails in folders on Microsoft Outlook. Enable macros to run include VBA Scripts in rules to download the attachments to specified folders in my Microsoft OneDrive account accessible from anywhere. Press Alt F11 to access Microsoft Visual Basic for Applications to set up VBA Scripts.
3)  Verify "WriteToReport.py" will copy and paste from the correct cells in your scheduled reports to the excel workbook you wish to copy to. 
______________________________________________________
Daily -
1)  Direct conversion python script to the correct folders. These scripts are set ip to convert onky the two mest files in the folder so it dorsnt modify ebrry filr im the folder changing thr modifoed dates listed. 
2)  Schedule "ScheduledPythonFiles.py" for the following morning prior to arriving to work (8:30 AM CST) and run the script. Several of the scripts convert each of the downloaded attachments from either .xls and .csv to .xlsx in order to work with the Openpyxl library in Python. The last script displays the newly converted file names for easy copying and pasting into "WriteToReport.py".
3)  Upon arriving to work, I copy and paste each of the newly converted file names into "WriteToReport.py". The names are listed on my Command Line Interface after scheduling and running "ScheduledPythonFiles.py" overnight.
4)  Run "WriteToReport.py". All data is copied to the excel report (from 6 different files in my case) and the file is saved to the OneDrive Directory / SharePoint I created for the team. Now my leadership team has access, and I didn't have to send one email. 



VOILA!
