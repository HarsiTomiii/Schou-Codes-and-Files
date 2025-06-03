import win32com.client
import os

# run this command in the console to export it into an executable
# pyinstaller --onefile --hidden-import win32com.client main.py

def run_excel_macro(file_path, macro_name):
    # Create Excel application object
    excel = win32com.client.Dispatch("Excel.Application")

    try:
        # Make Excel visible (set to False if you want it to run in the background)
        excel.Visible = True

        # Open the workbook
        print(f"Opening workbook: {file_path}")
        workbook = excel.Workbooks.Open(file_path)

        # Run the macro
        print(f"Running macro: {macro_name}")
        excel.Application.Run(f"'{os.path.basename(file_path)}'!{macro_name}")

        # Save and close the workbook
        workbook.Close(SaveChanges=True)
        print("Workbook closed successfully")

    except Exception as e:
        print(f"Error: {str(e)}")

    finally:
        # Quit Excel application
        excel.Quit()
        # Release COM objects
        del excel


if __name__ == "__main__":
    # File path
    file_path = r"C:\Users\tamhar\Schou Company A S\SC - INTERNT SUPPLY - Documents\General\Datafiles\Warehouses\Schou - Vision Park\warehouse_empty_times.xlsm"

    # Macro name
    macro_name = "refresh_all"

    # Run the function
    run_excel_macro(file_path, macro_name)