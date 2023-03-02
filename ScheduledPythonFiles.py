import schedule
import time
import subprocess


def run_files():
    # run the first python file
    subprocess.run(["python", "C:\\Users\\Cprickett\\Python-File-To-Convert-XLS-To-XLSX"])

# run the second python file
    subprocess.run(["python", "C:\\Users\\Cprickett\\Python-File-To-Convert-XLS-To-XLSX"])

# run the third python file
    subprocess.run(["python", "C:\\Users\\Cprickett\\Python-File-To-Convert-CSV-To-XLSX"])
    
# run the fourth python file
    subprocess.run(["python", "C:\\Users\\Cprickett\\NewFiles.py"])

#Schedule the code to run at 8:30 AM CST
schedule.every().day.at("08:30").do(run_files)

while True:
    schedule.run_pending()
    time.sleep(1)