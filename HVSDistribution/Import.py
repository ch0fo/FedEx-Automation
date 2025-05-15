
""" Libraries """
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import openpyxl
import logging
import logging.handlers
import datetime
import dotenv, os

def get_date():
    import datetime
    today = datetime.date.today().strftime("%d %b %Y")
    return today

""" Error Functions """
def NoFileError(): 
    noFile = messagebox.showerror("HVS Distribution", "No file was selected.")

def NoSheetError(): 
    noSheet = messagebox.showerror("HVS Distribution", "No sheet is of the correct format to import distribution.")

def OpenFileError():
    openFile = messagebox.showerror("HVS Distribution", "Please close the file before importing.")

def getFile():
    file_path = filedialog.askopenfilename(title="Select a File", filetypes=[("Excel file", "*.xlsx *.xlsm")])
    print("filepath: {}".format(file_path))
    return file_path

def getDistribution(label: tk.Label, stringvar: tk.StringVar):
    global file_path
    try:
        file_path = getFile()
        if file_path == "": 
            NoFileError()
            stringvar.set("No file selected!")
            label.config(fg='red')
            return False
        else: 
            # Load the Excel file
            wb = openpyxl.load_workbook(file_path)
            
            # Dictionary to store distribution of numbers
            global distribution
            distribution = {}
            
            # Flag to track if any sheet has correct headers
            correct_headers_found = False

            # Iterate through all sheets in the workbook
            for sheet in wb.worksheets:
                # Check if current sheet has been hidden by the user
                if sheet.sheet_state == "hidden":
                    print("WARNING: Sheet '{}' is hidden and will not be considered for data extraction.\n".format(sheet.title))
                    continue
                # Check if headers are correct
                headers = [cell.value for cell in sheet[1]]  # First row
                print(sheet[1])
                print("cells:")
                print(headers)
                print("done\n")

                # Checking number of headers
                if len(headers) < 24:
                    print("WARNING: Sheet '{}' has been skipped since it does not have enough headers.\n".format(sheet.title))
                    continue

                if headers[0] != "F1:AWB/TRACKING" or headers[10] != "Assign To..." or headers[23] != "EE_NME":
                    continue  # Skip to the next sheet if headers are not correct

                correct_headers_found = True  # Set the flag if correct headers are found

                # Iterate through rows
                for row in sheet.iter_rows(min_row=2, values_only=True):  # Assuming headers are in the first row
                    if not row[0] or not row[10] or not row[23]:
                        continue  # Skip the row if any of the required cells are blank

                    number = row[0]  # Value from column A
                    value = str(row[10]) + " - " + str(row[23])  # Value from column K and X
                    
                    # Check if the value exists in distribution dictionary
                    if value in distribution:
                        distribution[value].append(number)
                    else:
                        distribution[value] = [number]

            if not correct_headers_found:
                NoSheetError()
                return False

            # Save the workbook with the new data
            try:
                wb.save(file_path)
            except: #will break if user does not close the excel file they're trying to save
                OpenFileError()
                stringvar.set("Please close the file before importing.")
                label.config(fg='red')
                file_path = "" #removing filepath, even if it loaded before, to avoid confusion in the main GUI
                return False
            
            print(distribution)
            print(f"Brokers to process: {len(distribution)}")
            time: str = datetime.datetime.now().strftime('%H:%M:%S')
            label.config(fg='forest green')
            stringvar.set(f"Import successful as of {time}")
            return True
    except Exception as e:
        print('Exception: ', e)
        user = ""
        #trying to find user for exception, when user provided their ID
        try:
            dotenv.load_dotenv(dotenv_path=dotenv.find_dotenv(), override=True)
            user = os.environ['okta_username']
        except:
            user = "unknown"