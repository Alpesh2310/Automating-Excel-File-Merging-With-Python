import pandas as pd
import os
from tkinter import Tk, filedialog
import time

# Hide the root window
Tk().withdraw()

# Ask user to select a folder
folder_path = filedialog.askdirectory(title="Select Folder Containing Excel Files")
if not folder_path:
    print("❌ No folder selected!")
    exit()

# List all Excel files in the folder
files = [f for f in os.listdir(folder_path) if f.endswith(('.xlsx', '.xls'))]

if not files:
    print("❌ No Excel files found in the selected folder!")
    exit()

# Create empty list to store data
all_data = []

# Loop through files and read each Excel
for file in files:
    file_path = os.path.join(folder_path, file)
    df = pd.read_excel(file_path)
    df["Source_File"] = file   # Optional: keep track of source file
    all_data.append(df)

# Concatenate all dataframes
merged_df = pd.concat(all_data, ignore_index=True)

# Save to one Excel file
output_file = os.path.join(folder_path, "Merged_All_Files.xlsx")
merged_df.to_excel(output_file, index=False)

print(f"✅ All {len(files)} files merged into: {output_file}")
print("Closing in 5 seconds...........", flush=True)
time.sleep(5)
