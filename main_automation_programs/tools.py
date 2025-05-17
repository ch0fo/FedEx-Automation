"""
Mostly contains auxiliary functions used in HVS Distribution tool -> OKTA automation, chromedriver installation, saving user data.
Contains some functions used for emailing out file attachments (I mostly use these emailing functions to automate reports -> trigger sending an email through power automate when an email from these tools is received in my email)
Also contains other miscellaneous tools used throughout programs.

The .env file in this directory also contains FedEx login info for okta automation and other applications where FedEx logins are needed.
"""

#Imports
import os, chrome_version, requests, zipfile, io, time, certifi, dotenv
from pathlib import Path
from typing import List, Literal
import zipfile
from zipfile import ZIP_DEFLATED
import pandas as pd, numpy as np
from pandas import DataFrame
import classify_db
from typing import Dict

#Selenium imports
from selenium import webdriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#Email imports
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

#Literals
_connector = Literal['AND', 'OR']

#Getting SSL certifications -> NOTE: This SSL certification stuff is usually not required for running the programs locally, but you might need to include it when
#                                    exporting your programs into an .exe or similar. I keep it commented out because it slows down running my programs locally
#                                    a lot. Take the comments off when exproting programs to executables.

#--- COMMENT THIS PART OFF WHEN EXPORTING PROGRAMS
# try:
#     ssl_path: str = str(Path(os.getcwd())/'certifi'/'cacert.pem')
#     assert os.path.exists(ssl_path), 'No SSL certificates found in the root directory.'
#     os.environ['REQUESTS_CA_BUNDLE'] =  ssl_path #include certificates in root of .exe instance, when exported
#     print('SSL certificates retrieved successfully.')
# except Exception as e:
#     print(f'Failed to retrieve SSL certificates. Exception: {e}')
#--- COMMENT THIS PART OFF WHEN EXPORTING PROGRAMS

def get_envvar(var_name: str, envfile_path: str = dotenv.find_dotenv()) -> str:
    """
    Returns the value for the given env var name.

    If the var does not exist, it returns ''.

    If no env file path is given, it will load the one closest to where the tools.py file is stored.
    """

    dotenv.load_dotenv(dotenv_path=envfile_path, override=True) #loading given env file

    return os.environ[var_name]

def get_chromedriver(env_path: str = '') -> str:
    """
    Will get saved path for manual chromedriver initalization.\n
    If it fails to automatically download for current chrome version, will prompt to enter an installation path manually.
    """

    driver_path: str = ''
    
    #Attempting to automatically install correct chromedriver version
    try:
        driver_path = auto_path_download()
    except Exception as e:
        print(f"Failed to automatically download chromedriver for current chrome version. Exception: {e}")
        driver_path = get_manual_chromedriver()

    return driver_path

def get_manual_chromedriver() -> str:
    """
    Gets chromedriver path manually from user.
    """
    #Checking if .env file exists, creating if
    # env_path: Path = Path(os.path.dirname(os.path.realpath(__file__)))
    # if not os.path.exists(env_path/'.env'):
    #     os.mkdir(env_path/'.env')

    dotenv.load_dotenv(dotenv_path=dotenv.find_dotenv(), override=True) #Automatically finds closest .env file

    driver_path: str = ''
    chrome_curr_version: str = chrome_version.get_chrome_version()

    try: #checking if path has already been provided
        driver_path = os.environ['chromedriverpath']
        print(f"Retrieving chromedriver path found in .env file: '{driver_path}'")

        #Will also need to check if version is outdated
        curr_version: str = os.environ['chromedriver_version'] #If the statement made it this far, there must always exist a saved chromedriver version, as it has been set before
        
        if curr_version != chrome_curr_version: #checking if saved version matches current chrome installation version
            raise Exception(f"The saved chromedriver version '{curr_version}' does not match the installed chrome browser version '{chrome_curr_version}'")
        #If versions do not match, raise exception so it can be reset

    except Exception as e: #if not provided/outdated, prompt for new chromedriver path
        print(f'Setting a new chromedriver path. Exception: {e}')
        driver_path = input("Please enter your chromedriver path in the format ' \"my\\path\\to\\chromedriver.exe\" ': ")#getting path from user
        driver_path = driver_path.replace('\"','') #removing quotations, if encoutered
    
        #saving path to .env
        dotenv.set_key(dotenv_path=dotenv.find_dotenv(), key_to_set='chromedriverpath', value_to_set=driver_path)

        #saving version to .env
        dotenv.set_key(dotenv_path=dotenv.find_dotenv(), key_to_set='chromedriver_version', value_to_set=chrome_curr_version)

    return driver_path

def auto_path_download() -> str:
    """
    Attempts to find version for current chrome installation, and attempts to download matching chromedriver. Returns path of download.
    """

    #Initial setup
    chrome_vrsn: str = chrome_version.get_chrome_version() #current chrome version
    print(f"Current chrome version: '{chrome_vrsn}'")

    #Downloading chromedriver, if needed (first time download/outdated)
    if new_download(chrome_vrsn):
        default_driverurl = r'https://storage.googleapis.com/chrome-for-testing-public/{version}/win64/chromedriver-win64.zip'
        print(f"Downloading from '{default_driverurl.format(version = chrome_vrsn)}'")
        req: requests.Response = requests.get(default_driverurl.format(version = chrome_vrsn)) #inserting version to download link
        zip_fl: zipfile.ZipFile = zipfile.ZipFile(io.BytesIO(req.content)) #creating zipfile object from request's response
        zip_fl.extractall() #extracting all contents to current working directory
    else:
        print(f'No new download needed, using current chromedriver for version {chrome_vrsn}')
    
    #Saving latest chromedriver version download to closets dotenv file
    dotenv.set_key(dotenv_path=dotenv.find_dotenv(), key_to_set='chromedriver_version', value_to_set=chrome_vrsn) #set to latest chrome version, same as latest downloaded
    print(f"Chrome version saved to records for comparison: {get_envvar('chromedriver_version')}")

    #getting download path
    driver_path: Path = Path(os.getcwd())/r'chromedriver-win64\chromedriver.exe'
    print(f"Using installed chromedriver  at '{str(driver_path)}'")

    return str(driver_path) #returning path

def new_download(curr_version: str) -> bool:
    """
    Checks if a new download is required for the chromedriver, given the current chrome browser version
    """

    installed_version: str = ''
    env_path: str = dotenv.find_dotenv()
    print(f"Now looking at env path {env_path}")
    dotenv.load_dotenv(dotenv_path=env_path, override=True) #overrides and loads latest dotenv variables from closest .env file
    #Attemps to get currently installed version from .env variables
    try:
        installed_version = os.environ['chromedriver_version'] #attempts to get saved driver version from .env variables
    except:
        print('No saved chromedriver version found.')
        return True #If no saved version found, need to download new drivers (since no drivers have been previously downloaded)
    
    #Checking if installed version matches latest browser version
    if curr_version == installed_version:
        return False #in case they match, no new download needed
    
    return True #In any other case, need new download (didn't match)

#These tools are used to automate okta logins
def okta_login(driver: webdriver.Chrome, target_xpath: str, username: str, pssword: str) -> None:
    """
    Automatically completes okta login.\n

    target_xpath: path to look for in the next page after okta; used to verify if target website has already loaded.\n
    username: username to use for okta login.\n
    pssword: password to use for okta login.
    """
    done: bool = False
    while not done:
            try:# Will first try to find the desired website, in case okta is already validated
                driver.find_element(By.XPATH, target_xpath) #finding initial website element
            except: #Unable to find element, will attemp to complete okta login
                try:
                    user: WebElement = driver.find_element(By.XPATH, """//*[@id="input28"]""") #okta username field
                except:
                    continue
                else: #If okta login found, take rest of elements and complete login
                    password: WebElement = driver.find_element(By.XPATH, """//*[@id="input36"]""")
                    submit: WebElement = driver.find_element(By.XPATH, """//*[@id="form20"]/div[2]/input""")
                    print("Completing Okta signin.")
                    completing_okta(driver= driver, user=user, password=password, submit=submit,
                                    target_xpath= target_xpath, username = username, pssword = pssword) #completing okta login
                    print("Signin done, waiting for site")
                    completed: bool = False
                    while not completed:
                        try: #Now, waiting for target site to load
                            driver.find_element(By.XPATH, target_xpath)
                        except:
                            continue
                        else:
                            print("Site loaded, processing")
                            completed = True
            else:
                done = True
                print("Okta login has been completed")

    return None

def completing_okta(driver: webdriver.Chrome, user: WebElement, password: WebElement, submit: WebElement, target_xpath: str, username: str, pssword: str) -> None:
    """
    Once an okta login has been identified, automates okta login.
    """
    user.send_keys(username) #pastes user login, as provided when initializing user login (skips first two chars)
    password.send_keys(pssword) #pastes user password, as provided when initializing user login
    submit.click() #clicks submit button

    done = False
    print("Waiting for Submit button.")
    while not done:
        try:
            send: WebElement = driver.find_element(By.XPATH, """//input[@class=\button button-primary"]""")#old send button
        except: #if unable to find old submit button, looks for new button
            print("Attempting to find new submit button")
            try:
                # WebDriverWait(driver, 3).until(EC.element_to_be_clickable#clicking send push button
                #                         ((By.XPATH, """//*[@id="form53"]/div[2]/div/div[2]/div[2]/div[2]/a"""))).click()
                newSubmit: WebElement = WebDriverWait(driver, timeout=3.0).until(
                        EC.presence_of_element_located((By.XPATH, """/html/body/div[2]/div[2]/main/div[2]/div/div/div[2]/form/div[2]/div/div[2]/div[2]/div[2]/a"""))
                )
                newSubmit.click()
                print("New button found")

            except:
                #Target page might have already loaded, try to find initial element for it
                try:
                    driver.find_element(By.XPATH, target_xpath)
                except:
                    continue
                else:
                    done = True #if page loaded, terminate loop
            else:
                done = True #if new submit button found and clicked, terminate loop
        else:
            done = True #if old send button fouind, click it and termiante loop
            send.click()

    return None

def get_userlogindatapath() -> str:
    """
    Returns path where user login data is saved.
    """
    path = Path(os.getcwd())/'userdata'
    if not Path.exists(path):
        os.mkdir(path)
    print(f"User data path: '{path}'")
    return str(path)

def list_to_str(list: list) -> str:
    """
    Converts given list into a string. Elements are separated by newlines.
    """

    string: str = ''
    for item in list:
        string += f'{item}\n'

    return string[:-1] #returning without the last newline character

def get_query(query_path: str) -> str:
    """
    Reads query as a string from given .sql file.
    """
    with open(query_path, 'r') as sql_file:
        query: str = sql_file.read()

    return query

def send_email(
        subject: str,
        files: List[str],
        msg: str = "",
        send_to: List[str] = [get_envvar('fedex-email')],
        send_from: str = get_envvar('gmail-email'),
        server: str = 'smtp.gmail.com',
        port: int = 465,
        credentials_username: str = get_envvar('gmail-email'),
        credentials_password: str = get_envvar('gmail-app-password'),
        compress_files: bool = False
    ) -> None:
    """
    Sends the given email, with the given parameters.

    Please provide 'files' a a list of paths in your computer to get files from.
    Note that it accepts multiple people to send to; provide list of email addresses to email to.

    This is mainly used to automate flows in power automate; whenever I get an email, it uploads it to sharepoint for me, etc.

    Compress: compresses the given files into a single .zip file; files are sent out zipped.
    """

    #Creating email message
    message: MIMEMultipart = MIMEMultipart()
    message['From'] = send_from
    message['To'] = ', '.join(send_to)
    message['Date'] = formatdate(localtime=True)
    message['Subject'] = subject

    message.attach(MIMEText(msg)) #Attaching main message

    #Checking if files need to be compressed
    if compress_files:
        files = zip_files(files)

    #Appending attachments
    for path in files: #I dont understand any of this code, I copied it from https://stackoverflow.com/questions/3362600/how-to-send-email-attachments and https://mailtrap.io/blog/python-send-email-gmail/
        attachmt = MIMEBase('application', 'octet-stream')
        with open(path, 'rb') as file:
            attachmt.set_payload(file.read())
        encoders.encode_base64(attachmt)
        attachmt.add_header('Content-Disposition',
                f'attachment; filename={Path(path).name}') #Adding filename
        message.attach(attachmt) #adding file attachment to email message

    with smtplib.SMTP_SSL(server, port) as smtp_server: #This seems to not require tls anymore
       smtp_server.login(credentials_username, credentials_password)
       smtp_server.sendmail(send_from, send_to, message.as_string())
    print(f"Message with subject '{subject}' sent to {send_to}!\n{len(files)} attachment(s) sent.")

    return None

def zip_files(to_zip: List[str], zip_name: str = 'files.zip') -> List[str]:
    """
    Zips given files into a single zip file.

    The .zip file will be saved in the local directory, and the path for it will be returned as a single element list.
    """

    #Creating and saving zip object with all the given files to root
    with zipfile.ZipFile(zip_name, mode="w", compression=ZIP_DEFLATED) as archive:
        for filename in to_zip:
            curr_path = Path(filename)
            archive.write(curr_path, str(curr_path).split('/')[-1])

        for zipped_file in archive.filelist:
            zipped_file.filename = zipped_file.filename.split('/')[-1]

    print(archive.filelist)

    return [str(Path(os.getcwd())/zip_name)]

def get_password() -> str:
    """
    Loads closest .env file and gets 'pw' from it.
    """

    dotenv.load_dotenv(dotenv_path=dotenv.find_dotenv(), override=True) #Loading closest oven to get latest FedEx password

    return os.environ['pw']

def merge_files(
        to_merge: List[str],
        dtypes: Dict[str, object] = {}, #use this dict to specify what datatypes you want to sue for your file's columns. See the dtypes_complete variable below for an example
        dates: List[int] = [],
        # complete: bool = False,
        awb_col_name: str = 'Tracking Number'
        ) -> pd.DataFrame:
    """
    Merges given files into a single dataframe

    to_merge: list of paths, as string list, to merge
    """
    print("Merging selected files into one dataframe")
    # headers_partial = {'Tracking Number': 'awb',
    #        'Shipment Value': 'cad_val',
    #        'Importer Name': 'importer'}
    # dtypes_complete = {'Tracking Number': np.int64,
    #        'Shipment Value': np.float64
    # }
    merged_df: pd.DataFrame = pd.DataFrame()

    #merging fetched files
    for path in to_merge:
        try:
            # if not complete:
            #     dataframe = pd.read_csv(filepath_or_buffer=path, usecols=['Tracking Number', 'Shipment Value', 'Importer Name'],
            #         dtype={'Tracking Number': np.int64, 'Shipment Value': np.float64, 'Importer Name': str})
            #     dataframe = dataframe.rename(columns=headers_partial)
            if dtypes:
                dataframe = pd.read_csv(filepath_or_buffer=path, dtype=dtypes, parse_dates=dates)
            else:
                dataframe = pd.read_csv(filepath_or_buffer=path, parse_dates=dates)
        except Exception as e: #may be unable to read a file properly; clean faulty lines
            if isinstance(e, ValueError):
                #Reading awb col to find OLD awbs
                dataframe = pd.read_csv(filepath_or_buffer=path, usecols=[awb_col_name])
                drop = list()
                for row in dataframe.itertuples():
                    try:
                        int(row._1)
                    except:
                        print(f"Faulty line: {row}")
                        drop.append(int(row.Index)) #Keeping track of all OLD awbs
                # reading again and skipping faulty lines
                # if not complete:
                #     dataframe = pd.read_csv(filepath_or_buffer=path, usecols=['Tracking Number', 'Shipment Value', 'Importer Name'],
                #         dtype={'Shipment Value': np.float64, 'Importer Name': str})
                # else:
                dataframe = pd.read_csv(filepath_or_buffer=path, parse_dates=dates)
                print("Dropping faulty")
                dataframe = dataframe.drop(drop, inplace=False)
                # if not complete:
                #     dataframe = dataframe.astype(dtypes_complete)
                #     dataframe = dataframe.rename(columns=headers_partial)
                # else:
                dataframe = dataframe.astype(dtypes)
                if merged_df.shape[0] == 0:
                    merged_df = dataframe.copy(deep=True)
                else:
                    merged_df = pd.concat([merged_df, dataframe], ignore_index=True)
                print(f'Merged after concat with fixed faulty:\n{merged_df}')
            else:
                print(f"Unable to read file '{path}', verify file structure\nException: {e}, exception type: {type(e)}")
        else:
            if merged_df.shape[0] == 0:
                merged_df = dataframe.copy(deep=True)
            else:
                merged_df = pd.concat([merged_df, dataframe], ignore_index=True)

    return merged_df

def sql_able_list(vals: list, logic: str, variable: str, connector: _connector = 'OR') -> str:
    """
    Converts given list of values into an sql compatible list (split into chunks of 1000)\n

    vals = list of values to parse as sql compatible\n
    logic = how to join list of values (e.g: IN, NOT IN, etc.)\n
    variable = name of relevant variable (e.g: AWB, AWB_NBR, etc.)\n
    connector = how to connect chunks ('AND', or 'OR'), defaults to 'OR'
    """
    chunksize: int = 1000
    n_chunks: int
    chunks = list()
    final_sql = ''

    for i in range(0, len(vals), chunksize):
        chunks.append(vals[i:i+chunksize])#pasting all vals as chunks of size chunksize

    n_chunks = len(chunks)
    for i in range(0, n_chunks, 1):
        curr_sql = ''
        len_curr = len(chunks[i])
        for j in range(0, len_curr, 1):
            curr_sql += f'\'{chunks[i][j]}\''
            if j != len_curr-1:
                curr_sql += ','
        temp = f'{variable} {logic} ({curr_sql})'
        if i != n_chunks-1:
            temp += f' {connector} '
        final_sql += temp
    return final_sql

def get_classify(awbs: list, query: str) -> pd.DataFrame:
    """
    Returns data for given awbs if found in Classify.\n
    Returns data with awb number as index in dataframe.
    """
    classify_data: pd.DataFrame = pd.DataFrame()
    chunksize = 65000
    chunks = []
    for i in range(0, len(awbs), chunksize):
        chunks.append(awbs[i:i+chunksize])

    #concating chunks
    n_chunks = len(chunks)
    for chunk in chunks:
        print(f"Classify data chunks left: {n_chunks}")
        classify_query = query.format(
            awbs_to_search = sql_able_list(vals=chunk, logic='IN', variable='AWB_NBR', connector='OR'))
        dataframe = classify_db.execute_query(query= classify_query)

        if classify_data.shape[0] == 0:
            classify_data = dataframe
        else:
            classify_data = pd.concat([classify_data, dataframe], ignore_index=True)
        n_chunks -= 1
    
    #dropping dups
    classify_data = classify_data.drop_duplicates(subset=['awb_nbr'])

    #changing index, returning
    return classify_data.set_index('awb_nbr', drop=True)

def find_OLD(dataframe: pd.DataFrame, awb_position: int = 1) -> List[int]:
    """
    From given dataframe, finds index of awbs that cannot be casted to int (usually because they contain 'OLD'). Returns list of indices.\n

    awb_position: in what position to look for awb value (most usually in second position, index one, after dataframe index)
    """

    drop: List[int] = list()

    faulty_count: int = 0

    for row in dataframe.itertuples():
        try:
            int(row[awb_position])
        except Exception as e:
            print(e, type(e))
            print(f"Faulty line: {row}")
            faulty_count += 1
            drop.append(int(row[0])) #if unsuccessful cast, append index so it can be dropped.

    print(f"Faulty dropped: {faulty_count}")
    return drop

def merge_files_df(files: List[str], clean_OLD: bool = True, awb_pos: int = 1) -> DataFrame:
    """
    Merges given files (defaults to merging csv files) into a single dataframe and returns it.
    By default, will attempt to find and remove all 'OLD' containing awbs it can find in the specified awb column index (defaults to second position, index 1).
    Assumes that .csv files are being passed.
    """

    merged: DataFrame = DataFrame()

    #Merging given files
    for file in files:
        try:
            dataframe: DataFrame = pd.read_csv(filepath_or_buffer=file)
        except Exception as e:
            print(f"Error reading file '{file}', exception: {e}")
        else:
            try:
                if merged.shape[0] == 0:
                    merged = dataframe.copy(deep=True) #copying first dataframe directly if master merged df is empty
                else:
                    merged = pd.concat([merged, dataframe], ignore_index=True) #else, appending to master merged df
            except Exception as e:
                print(f"Unable to merge file '{file}', exception: {e}")

    #Cleaning 'OLD' awbs, when applicable:
    if clean_OLD: merged = merged.drop(find_OLD(merged, awb_position= awb_pos), inplace=False)

    return merged