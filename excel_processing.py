import pandas as pd
import os

def format_dataframe(df):
    # Convert all columns to string for uniform formatting
    df = df.astype(str)
    
    # Adjust the column widths
    df.iloc[:, 0] = df.iloc[:, 0].str.zfill(16).str.ljust(17)  # First column: width of 17, preserve leading zeros
    df.iloc[:, 1] = df.iloc[:, 1].str.ljust(8)   # Second column: width of 8
    df.iloc[:, 2] = df.iloc[:, 2].str.ljust(4)   # Third column: width of 4
    df.iloc[:, 3] = df.iloc[:, 3].str.ljust(4)   # Fourth column: width of 4

    return df

def excel_to_prn(excel_file_path, output_directory, identifier):
    try:
        # Read the Excel file
        excel_data = pd.ExcelFile(excel_file_path)
        
        # Count the number of sheets
        number_of_sheets = len(excel_data.sheet_names)
        print(f"The Excel file contains {number_of_sheets} sheet(s).")
        
        # Ensure the output directory exists
        os.makedirs(output_directory, exist_ok=True)

        # Loop through each sheet in the Excel file
        for sheet_name in excel_data.sheet_names:
            # Skip the sheet named "Summary"
            if sheet_name.lower() == "summary":
                print(f"Skipping sheet '{sheet_name}'.")
                continue

            # Read each sheet into a DataFrame
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

            # Delete the first row
            df = df.iloc[1:].reset_index(drop=True)

            # Remove empty rows
            df = df.dropna(how='all')

            # Apply formatting
            formatted_df = format_dataframe(df)

            # Display the first row of the formatted DataFrame in the console
            print(f"First row of formatted data from sheet '{sheet_name}':")
            print(formatted_df.head(1).to_string(index=False, header=False))

            # Prepare the custom PRN file name
            prn_file_name = f"TGPAP.GPNMA.{identifier}.CT{sheet_name}.prn"
            prn_file_path = os.path.join(output_directory, prn_file_name)
            
            # Save formatted DataFrame to a .prn file with spaces as the delimiter
            with open(prn_file_path, 'w') as file:
                # Join all columns into a single string for each row
                lines = formatted_df.apply(lambda row: ''.join(row), axis=1).tolist()
                # Write lines to the file, avoiding an extra newline at the end
                file.write('\n'.join(lines))

            # Notify that the file has been created
            print(f"File '{prn_file_path}' has been created successfully.")

    except Exception as e:
        print(f"An error occurred: {e}")

# Example usage
excel_file_path = r'C:\Users\AlexandruMarin\Downloads\20240916_PH.xlsx'  # Excel file path
output_directory = r'C:\Users\AlexandruMarin\Downloads'  # Output directory to save the PRN files
identifier = 'AMN.F49718'  # Variable part of the filename
excel_to_prn(excel_file_path, output_directory, identifier)