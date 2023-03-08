import os
import pandas as pd
from datetime import datetime

# Specify the path of the folder containing CSV files
csv_folder = r'C:\\Users\\Cprickett\\INPUT-FOLDER-WITH-CSV-FILES'

# Specify the date on which the CSV files were modified (in YYYY-MM-DD format)
specified_date = '2023-03-02'  # Change the date here

def convert_csv_to_xlsx():
    # Get a list of all CSV files modified on the specified date in the folder
    csv_files = [f for f in os.listdir(csv_folder) if f.endswith('.csv') 
                 and datetime.fromtimestamp(os.path.getmtime(os.path.join(csv_folder, f))).date() == datetime.strptime(specified_date, '%Y-%m-%d').date()]

    # Loop through each CSV file and convert it to XLSX format
    for csv_file in csv_files:
        # Load the CSV file into a pandas dataframe
        df = pd.read_csv(os.path.join(csv_folder, csv_file))

        # Create a new filename for the XLSX file
        xlsx_file = os.path.splitext(csv_file)[0] + '.xlsx'

        # Save the dataframe to the new XLSX file
        df.to_excel(os.path.join(csv_folder, xlsx_file), index=False)
        
        print(f"Converted {csv_file} to {xlsx_file}")

convert_csv_to_xlsx()

print('Blackboard conversion complete')