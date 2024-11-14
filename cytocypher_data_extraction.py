#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
from pathlib import Path

# Define the input file path
input_file_path = Path(r"C:\Users\User\Desktop\example.xlsx")
# Define the output folder path
output_folder_path = Path(r"C:\Users\User\Desktop\")

# Check if the input file exists
if not input_file_path.exists():
    print(f"Error: File not found at {input_file_path}")
else:
    # Read the Excel file
    xlsx = pd.ExcelFile(input_file_path)

    # Create an empty list to store the data
    all_data = []

    # Loop through each sheet in the Excel file
    for sheet_name in xlsx.sheet_names:
        # Read the sheet data into a DataFrame
        df = xlsx.parse(sheet_name)

        # Check if the required column names exist in the first row
        if 'Peak' in df.columns and 'Departure velocity' in df.columns and 'Return velocity' in df.columns and 'Percentage of shortening' in df.columns:
            # Extract the desired columns based on their column names
            peak_col = df['Peak height']
            time_to_peak_90_col = df['Departure velocity']
            time_to_baseline_90_col = df['Return velocity']
            percentage_of_shortening_col = df['Percentage of shortening']

            # Add the sheet name as the first row
            sheet_data = [f"Sheet name: {sheet_name}"]
            sheet_data.append("Peak height\t" + "\t".join(map(str, peak_col.tolist())))
            sheet_data.append("Departure velocity\t" + "\t".join(map(str, time_to_peak_90_col.tolist())))
            sheet_data.append("Return velocity\t" + "\t".join(map(str, time_to_baseline_90_col.tolist())))
            sheet_data.append("Percentage of shortening\t" + "\t".join(map(str, percentage_of_shortening_col.tolist())))

            # Add the sheet data to the list
            all_data.extend(sheet_data)
        else:
            print(f"Warning: Required column names not found in sheet '{sheet_name}'. Skipping this sheet.")

    # Write the result to a new text file
    output_file_path = output_folder_path / "combined_data.txt"
    with open(output_file_path, 'w') as file:
        for data in all_data:
            file.write(data + "\n")

