import pandas as pd

# File path for the uploaded Excel file
excel_file_path = r"\\hc22itsup\zCache\Service Desk\Tools\LicenseMon\Ref\Microsoft Licenses Requested FY 24 8.31.2023.xlsx"

# Load the specific workbook and the required columns and rows
df = pd.read_excel(excel_file_path, sheet_name='Summary Requested FY 24 O365', usecols='O,P,Q,R', nrows=21)

# Define the path for the output CSV file
output_csv_path = r"\\hc22itsup\zCache\Service Desk\Tools\LicenseMon\Ref\cellPull.csv"


# Save the dataframe to a CSV file
df.to_csv(output_csv_path, index=False)

print("pulled maxLic")
