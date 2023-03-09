import os
import glob

folders = [r'C:\\LOCATION-FOR-NEWLY-CONVERTED-FILES\\',
    r'C:\\LOCATION-FOR-NEWLY-CONVERTED-FILES\\',
    r'C:\\LOCATION-FOR-NEWLY-CONVERTED-FILES\\']

for folder in folders:
    os.chdir(folder)
    files = sorted(glob.glob("*"), key=os.path.getctime, reverse=True)[:2]
    print(f"Latest files in {folder}: {files}")
