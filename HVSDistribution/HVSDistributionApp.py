
""" Files """
import Import
import Upload
import customCheckbox

""" Libraries """
import tkinter as tk
import os
import sys
import dotenv
import datetime

""" Utilities """
def initializeFlag(): 
    global distributionFlag
    distributionFlag = False
    global notShowAgainFlag
    notShowAgainFlag = False

def displayFileName(): 
    if Import.file_path != "": 
        file_name = os.path.basename(Import.file_path)
        selectedFileLabel.config(text= f"Selected File: {file_name}") 
    else: 
        return
    
def distribute(message: tk.Label, text_var: tk.StringVar): 
    """
    The 'message' label is used to update the user on the status of the file import.
    The 'text_var' string var is used to change the text 
    """
    global distributionFlag
    distributionFlag = Import.getDistribution(message, text_var)

def upload(): 
    if distributionFlag == True: 
        global notShowAgainFlag
        if notShowAgainFlag == False: 
            okFlag, notShowAgainFlag = customCheckbox.show_custom_warning(root)
            # print("Dialog result:", okFlag)
            # print("Do not show again:", notShowAgainFlag)
            if okFlag == True: 
                Upload.setDistribution(width, height)
                return
            else: 
                return
        else: 
            Upload.setDistribution(width, height)
            return
    else: 
        Import.NoFileError()
        return
    
def verify_path(check_path: str) -> None:
    check_path = check_path.replace('\"', '')
    if os.path.exists(path=check_path):
        dotenv.set_key(dotenv_path=dotenv.find_dotenv(), key_to_set= 'chromedriverpath', value_to_set=check_path)
        print(f"Path '{check_path}' added to env")
    else:
        print("WARNING: Could not find given chromedriver path, env. var was not modified")

def get_credentials() -> tuple[str, str]:
    """
    Attempts to get credentials saved in .env file. If no credentials set, displays default values of 'Fedex ID' and '' (empty for password).
    """

    #Loading environ variables from closest .env file (ideally placed at root dir)
    dotenv.load_dotenv(dotenv_path=dotenv.find_dotenv())

    username: str = ''
    password: str = ''

    #Attempting to get username
    try:
        username = os.environ['okta_username']
    except:
        print('No username configuration set')
        username = 'Fedex ID'

    #Attempting to get password
    try:
        password = os.environ['okta_password']
    except:
        print('No password configuration set')
        password = ''

    return username, password


def update_credentials(username: str, password: str) -> None:
    """
    Saves updated credentials to .env file, for easier login.
    """

    #Finding closest .env path
    env_path: str = dotenv.find_dotenv()

    #Saving username credentials; finds closest .env file to store at
    dotenv.set_key(dotenv_path=env_path, key_to_set='okta_username', value_to_set=username)

    #Saving password credentials
    dotenv.set_key(dotenv_path=env_path, key_to_set='okta_password', value_to_set=password)

    return None

def update_confirmation(string_var: tk.StringVar, confirmation: tk.Label) -> None:
    """
    Updates the credentials' confirmation after each change. Takes in text widget to update.
    """

    #Getting current time
    time: str = datetime.datetime.now().strftime('%H:%M:%S')

    string_var.set(f'Successfully updated as of {time}')
    confirmation.config(fg='forest green')

    return None

def update_okta() -> None:
    """
    Updates / adds env var showing if the user would like to automate their okta login or not.
    """
    dotenv.set_key(dotenv_path=dotenv.find_dotenv(), key_to_set='okta_automation', value_to_set=okta_check_int.get().__str__())

    return None

def init_okta_val() -> int:
    """
    Reads last settings set for okta automation and returns it.
    """
    dotenv.load_dotenv(dotenv_path=dotenv.find_dotenv(),override=True)
    return int(os.environ['okta_automation'])

"""Reading args"""
args = sys.argv
if len(args) > 1:
    verify_path(args[1])

""" UI """
root = tk.Tk()
root.title(" FedEx Express HVS Distribution App")
frame = tk.Frame(root)

width = root.winfo_screenwidth()
height = root.winfo_screenheight()
initializeFlag()

""" String vars """
update_string: tk.StringVar = tk.StringVar()

""" Title """
nameLbl = tk.Label(frame, text = "Classify HVS Distribution App", font=("TkDefaultFont", 15))
# nameLbl.grid(row=0, column=0, sticky="W")
nameLbl.pack()

""" Explanation """
reminderMsg1 = """Choose a file to extract the HVS distribution from. 
NOTE: The file must be closed before it can be opened here."""
reminderLabel1 = tk.Label(frame, text= reminderMsg1)
# reminderLabel.grid(row= 9, column= 0, pady= 5, padx = 10, sticky = "W")
reminderLabel1.pack()

"""Import Button"""
importBtn = tk.Button(frame, text = "Import", command= lambda: [distribute(confirmation, update_string), displayFileName()], font=("TkDefaultFont", 12))
# importBtn.grid(row=8, column=0, pady= 5, padx = 10, sticky = "W")
importBtn.pack()

""" File Imported Confirmation """
selectedFileLabel = tk.Label(frame, text="Selected File:")
selectedFileLabel.pack()

""" Explanation """
reminderMsg2 = """Upload the HVS distribution to Classify"""
reminderLabel2 = tk.Label(frame, text= reminderMsg2)
# reminderLabel.grid(row= 9, column= 0, pady= 5, padx = 10, sticky = "W")
reminderLabel2.pack()

""" Upload to Classify """
uploadBtn = tk.Button(frame, text= "Upload", command= lambda: [upload()], font=("TkDefaultFont", 12))    
# uploadBtn.grid(row=10, column=0, pady= 5, padx = 10, sticky = "W")
uploadBtn.pack()

""" Username input """
entry_label: tk.Label = tk.Label(frame, text="Please enter your Fedex ID and Password below")
entry_label.pack()
username_entry: tk.Entry = tk.Entry(frame, bd =3)
username_entry.insert(0, get_credentials()[0])
username_entry.pack()

""" Password input """
password_entry: tk.Entry = tk.Entry(frame, bd =3, show='*')
password_entry.insert(0, get_credentials()[1])
password_entry.pack()

""" Update credentials button """
credentials_button: tk.Button = tk.Button(frame, text='Update credentials',
                                          command= lambda: [update_credentials(username = username_entry.get(), password = password_entry.get()),
                                                            update_confirmation(string_var = update_string, confirmation=confirmation)],
                                          font=("TkDefaultFont", 12))
credentials_button.pack()

"""Automate Okta Options"""
okta_check_int = tk.IntVar(value=init_okta_val())
okta_check = tk.Checkbutton(frame, text="Automate Okta Login", variable=okta_check_int, 
                             onvalue=1, offvalue=0, command=update_okta)
okta_check.pack()

""" Credentials update confirmation """
confirmation: tk.Label = tk.Label(master=frame, textvariable=update_string, anchor= 'center', bd=3, font=("TkDefaultFont", 8), fg='forest green')
confirmation.pack()

frame.pack(padx=10, pady=10)
root.mainloop()