# Cytocypher Data Extraction Script

This repository contains a Python script for extracting key metrics from Cytocypher “individual transients” .xlsx files. The script identifies specific columns for analysis and outputs the data in a combined text file.

## Features
- Reads data from multiple sheets in an Excel file.
- Extracts the following metrics:
  - **Peak Height**
  - **Departure Velocity**
  - **Return Velocity**
  - **Percentage of Shortening**
- Outputs the extracted data into a `.txt` file, organized by sheet name.

## Requirements
- **Python 3.x**
- Required Python libraries: `pandas`, `pathlib`

## Usage

1. **Place the Excel File**: Save your .xlsx file in the path specified in `input_file_path`. (You can change this path in the script.)
2. **Set the Output Folder**: Specify the folder path where the output file will be saved in `output_folder_path`.
3. **Run the Script**: Execute the script with the following command:

## Output: 

A text file combined_data.txt will be generated in the specified output folder containing the extracted data. You can copy these data in a new Excel file using CTRL+A and CTRL+C and make further analysis in this Excel file. 

## Notes
1. Ensure that the Excel file has the specified column names: 'Peak height', 'Departure velocity', 'Return velocity', and 'Percentage of shortening'.
2. If any of these columns are missing from a sheet, the script will skip that sheet and print a warning message. Normally, you will see three messages when the code processes sheets with average data. 
3. Do not forget to define the input and output file path.
