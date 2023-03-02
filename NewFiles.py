import os
import glob

"""This python code was created to generate the newly converted file names 
to be copied and pasted into the WriteToReport.py code to automate the new report
using the data from the newly converted spreadsheet files"""

folders = [r'C:\\Users\\Cprickett\\OneDrive - SODEXO\\Blackboard\\',
    r'C:\\Users\\Cprickett\\OneDrive - SODEXO\\Labor\\LaborReports\\ActualvsSched\\',
    r'C:\\Users\\Cprickett\\OneDrive - SODEXO\\Labor\\LaborReports\\Punch%\\']

for folder in folders:
    os.chdir(folder)
    files = sorted(glob.glob("*"), key=os.path.getctime, reverse=True)[:2]
    print(f"Latest files in {folder}: {files}")
