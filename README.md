# AutomateReporting

This is my workflow for automating reading and writing to a daily excel report in my current role as a food service Retail Manager.



One time set up -
1)  I set up financial and labor excel reports to be emailed to me overnight. (Kronos Workforce Central & Blackboard) 
2)  Set up Microsoft rules to organize emails in folders on Microsoft Outlook and include VBA Scripts to download the attachments to specified folders in my Microsoft OneDrive account accessible from anywhere. 

Daily -
1)  Set the run time for "ScheduledPythonFiles.py" for the following morning prior to arriving to work (8:30 AM CST) and run the script. Several of the scripts convert each of the downloaded attachments from either .xls to .xlsx or .csv to .xlsx in order to work with the Openpyxl library in Python. The last script displays the newly converted file names for easy copying and pasting.  
2)  Upon arriving to work, I copy and paste each of the newly converted file names into WriteToReport.py
3)  Verify the data is in the correct cells in the files. Set up the cells for which you want to copy and write to. 
4)  Run "WriteToReport.py"



VOILA!
