# Cytocypher Data Extraction Script

This repository contains a Python script for extracting key metrics from Ion-Optix Cytocypher “individual transients” .xlsx files. The script identifies specific columns for analysis and outputs the data in a combined text file.

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
4. You can delete/change/add analysed parameters based on your needs (eg Time to Peak or Time to Base 90).

# **Postprocessing Steps for Excel File**
## **Initial Check for Outliers:**

1. Open the Excel file generated from the text file.
2. Inspect the columns containing the extracted metrics: 'Peak height', 'Departure velocity', 'Return velocity', and 'Percentage of shortening'.
3. Identify and remove outliers or erroneous data, especially from the last column (as it is often an outlier).
4. Your rejected data will be indicated as "NaN". **Note**: "NaN" is not taken into account in the "Average formula", but if you average only "NaN", there will be "#DIV/0!", which is acceptable and will not be used in subsequent analyses since when you copy it and insert in GraphPad it will be just an empty cell.

## **Exclude Pixel Intensity Rows:**
Delete all rows related to EdgeLength.

## **Average Data in Relevant Columns:**
Average the remaining values for each metric after removing outliers in column M (RatioMetricCalcium and SarcomereLength).

## **Transfer Processed Data to Sheet2:**
Copy columns A (Segment Names) and M (Processed Average Data) related to the RatioMetricCalcium from Sheet1 to Sheet2.
Insert cells with an INDIRECT formula from "example" file for data referencing. 
Copy columns A (Segment Names) and M (Processed Average Data) related to the SarcomereLength from Sheet1 to Sheet3.
Insert cells with an INDIRECT formula from "example" file for data referencing. 

## **Organize Data for Final Analysis:**

1. Transfer the processed data from Sheet2 and Sheet3 to the Sheet4 and copy headings from the corresponding Sheet4 from the "example" file or make it by yourself. 
2. Use the Cytosolver app to identify the segment(s) where measurements begin with cell №1. Note that the data is divided into separate runs, and after each run, the measurements restart from cell №1. Ensure you correctly split the data based on these runs.
E.G. For the provided "example" file, 
3. I'm using distinct colours to differentiate between calcium and shortening data.

## **Export for GraphPad:**
Copy your final data to the GraphPad and plot graphs based on your needs.
