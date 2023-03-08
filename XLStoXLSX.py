import pandas as pd
import os
import glob


def convert_xls_to_xlsx():
    folder_path = "C:\\Users\\Cprickett\\INPUT-FOLDER-WITH-XLSX-FILES\\"
    xls_files = glob.glob(os.path.join(folder_path, "*.xls"))
    xls_files_sorted = sorted(xls_files, key=os.path.getctime, reverse=True)[:2] # get the two most recent .xls files
    for file_path in xls_files_sorted:
        df = pd.read_excel(file_path)
        xlsx_file_path = file_path.replace(".xls", ".xlsx")
        df.to_excel(xlsx_file_path, index=False)
        print(f"Converted {file_path} to {xlsx_file_path}.")


convert_xls_to_xlsx()

print('Conversion complete!')