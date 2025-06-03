import win32com.client
import os
import datetime


# run this command in the console to export it into an executable
# pyinstaller --onefile --name maersk_forecast_update --hidden-import win32com.client main.py
# pyinstaller --onefile --name maersk_forecast_update --hidden-import win32com --hidden-import win32com.client --hidden-import win32com.client.dynamic --hidden-import win32timezone --hidden-import pythoncom --hidden-import pywintypes main.py

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

        # Generate timestamp
        now = datetime.datetime.now()
        timestamp = now.strftime("%Y%m%d%H%M%S")

        # Create save path with timestamp
        save_dir = r"C:\Users\tamhar\Schou Company A S\SC - INTERNT SUPPLY - Documents\General\Datafiles\container_planning_data"
        save_filename = f"{timestamp}_container_plan_raw.xlsx"
        save_path = os.path.join(save_dir, save_filename)

        # Disable alerts to bypass the macro warning
        excel.DisplayAlerts = False

        # Save as xlsx file with timestamp filename
        print(f"Saving file as: {save_path}")
        workbook.SaveAs(save_path, FileFormat=51)  # 51 is the file format code for .xlsx

        # Re-enable alerts
        excel.DisplayAlerts = True

        # Close the workbook
        workbook.Close(SaveChanges=False)  # Already saved with SaveAs
        print("Workbook closed successfully")
        print(f"File saved successfully as: {save_filename}")

    except Exception as e:
        print(f"Error: {str(e)}")

    finally:
        # Make sure alerts are enabled before quitting
        excel.DisplayAlerts = True
        # Quit Excel application
        excel.Quit()
        # Release COM objects
        del excel


if __name__ == "__main__":
    # File path
    file_path = r"C:\Users\tamhar\Schou Company A S\SC - INTERNT SUPPLY - Documents\General\weekly_container_planning.xlsm"

    # Macro name
    macro_name = "refresh_all"

    # Run the function
    run_excel_macro(file_path, macro_name)