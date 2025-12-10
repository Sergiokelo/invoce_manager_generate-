import pandas as pd

# Corrected file paths using raw strings
input_file_path = r"D:\doc\traçabilité\MNH\data_C3\TABLE D'ATTRIBUT MNH (1).xlsx"
output_file_path = r"D:\doc\traçabilité\MNH\data_C3\MHN_data_C3.xlsx"

# Load the data from the Excel file
df = pd.read_excel(input_file_path, sheet_name='Feuil1')  # Specify the sheet if necessary

# Remove suffixes like '-P1', '-P2', etc., from the "CODE" column
df['CODE'] = df['CODE'].str.replace(r'-P\d+', '', regex=True)

# Save the modified data to a new Excel file
df.to_excel(output_file_path, index=False)

print("File saved successfully to:", output_file_path)
