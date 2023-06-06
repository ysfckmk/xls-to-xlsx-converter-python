import os
import glob
import pandas as pd
import subprocess

# Function to kill a running process by name
def kill_process(process_name):
    subprocess.call(["taskkill", "/f", "/im", process_name])

# Function to convert XLS files to XLSX format
def convert_xls_to_xlsx(folder_path):
    kill_process("excel.exe")  # Kill any running Excel processes
    os.chdir(folder_path)  # Change directory to the specified folder path
    for file in glob.glob("*.xls"):  # Loop through each XLS file in the folder
        xls_file = pd.ExcelFile(file, engine='xlrd')  # Read the XLS file using xlrd engine
        sheet_names = xls_file.sheet_names  # Get the sheet names
        
        writer = pd.ExcelWriter(file[:-4] + ".xlsx", engine='openpyxl')  # Create an Excel writer with xlsx engine
        
        # Iterate through each sheet and write it to the new XLSX file
        for sheet_name in sheet_names:
            df = xls_file.parse(sheet_name)  # Read the sheet as a DataFrame
            df.to_excel(writer, sheet_name=sheet_name, index=False)  # Write the DataFrame to the sheet
        
        writer.save()  # Save the workbook
        writer.close()  # Close the writer

        print(f"{file} success")  # Print success message for each converted file

folder_path = "your-folder-path"  # Enter your folder path here
convert_xls_to_xlsx(folder_path)  # Call the function to convert XLS files to XLSX format
