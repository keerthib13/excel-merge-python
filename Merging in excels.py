import os
import pandas as pd

data_file_folder = r"C:\Users\Keerthi\OneDrive\Documents\python new\pyhton fita\Freelancing"

# List only .xlsx and .xls files
excel_files = [
    f for f in os.listdir(data_file_folder)
    if f.endswith(('.xlsx', '.xls'))
]

if not excel_files:
    print("⚠️ No Excel files found to merge.")
else:
    all_data = []
    for file in excel_files:
        print(f"📂 Loading {file} ...")
        df = pd.read_excel(os.path.join(data_file_folder, file))
        all_data.append(df)
    merged = pd.concat(all_data, ignore_index=True)
    merged.to_excel(os.path.join(data_file_folder, "merged_data.xlsx"), index=False)
    print("✅ Merge completed. Saved as merged_data.xlsx.")
