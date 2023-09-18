import xlrd
import openpyxl
import os
import subprocess

# Check if openpyxl is installed, and if not, install it
try:
    import openpyxl
except ImportError:
    print("openpyxl is not installed. Installing...")
    subprocess.run(["pip", "install", "openpyxl"])

# Check if xlrd is installed, and if not, install it
try:
    import xlrd
except ImportError:
    print("xlrd is not installed. Installing...")
    subprocess.run(["pip", "install", "xlrd"])

def convert_old_excel_to_modern():
    try:
        # Prompt the user for the directory and file name of the old Excel file
        old_excel_file = input("Enter the path to the old Excel file (e.g., /path/to/old_file.xls): ")

        # Verify that the specified file exists
        if not os.path.exists(old_excel_file):
            print("Error: The specified file does not exist.")
            return

        # Create a new modern Excel file
        modern_excel_file = "modern_file.xlsx"  # Replace with the desired path for the modern Excel file

        # Load the old Excel file using xlrd
        old_workbook = xlrd.open_workbook(old_excel_file)

        # Create a new modern Excel file
        modern_workbook = openpyxl.Workbook()

        # Copy the data from the old workbook to the new workbook
        for sheet_name in old_workbook.sheet_names():
            old_sheet = old_workbook.sheet_by_name(sheet_name)
            modern_sheet = modern_workbook.create_sheet(title=sheet_name)

            for row_num in range(old_sheet.nrows):
                row_data = old_sheet.row_values(row_num)
                modern_sheet.append(row_data)

        # Save the new modern Excel file
        modern_workbook.save(modern_excel_file)

        print(f"Conversion completed. Saved as {modern_excel_file}")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    convert_old_excel_to_modern()
