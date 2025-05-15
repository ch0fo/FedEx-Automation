import datetime
import wfm_db
import csv
import os
from pathlib import Path
import subprocess
import pandas as pd
import numpy as np
from typing import List
from tools import send_email, get_envvar, get_query
from smtplib import SMTPSenderRefused

def main(query_path: str = r'main_automation_programs\support-files\queries\wfm_reserve.sql', email_to: List[str] = [get_envvar('fedex-email')]) -> None:
    """
    Gets today's wfm reserve file from oracle db.

    Emails resuting .csv file to given addresses.

    Note that if .csv file is too large to send (over 25 MB), it will be compressed and reattemped. If reattemp fails, no file will be emailed.

    query_path: path to look for wfm query (in .sql format)
    email: email addresses to send results to, as list of strings
    """

    path = Path(os.path.expanduser('~')) / 'Downloads'
    tday = datetime.date.today().strftime("%d%b%y")
    path /= f'WFM_Reserve {tday}.csv'
    print(f"Downloading to '{path}'")
    wfm_data = wfm_db.execute_query(query=get_query(query_path)) #getting query results
    wfm_data.columns = [x.upper() for x in wfm_data.columns]#changing col names

    #Defining cols we would like to force as ints, iterating
    to_int: List[str] = ['BILL_TO_ACC', 'PCS', 'TRANSACTION_NBR', 'ASSIGNED_TO', 'REASON_CD', 'SHIPPER_ACCT']

    # Fixing to ints
    for header in to_int:
        wfm_data = int_data(wfm_data, header)

    print(f"reserve df:\n{wfm_data}")
    wfm_data.to_csv(path_or_buf=str(path),index=False, quoting=csv.QUOTE_STRINGS, date_format=r"%d-%b-%Y %H.%M.%S.%f")
    
    # Email reserve out
    if len(email_to) != 0:
        try:
            send_email(subject=f'WFM_RESERVE_0001 {datetime.datetime.now().strftime("%d-%b-%Y")}',
                msg='', files=[str(path)], send_to=email_to)
        except SMTPSenderRefused as e:
            print(f"Failed to send email, most likely due to file exceeding size limit (can only send up to 25 MB).\n\n{e}\n\nReattempting with zipped file.")

            #Reattempting send with zipped file instead
            try:
                send_email(subject=f'WFM_RESERVE_0001 {datetime.datetime.now().strftime("%d-%b-%Y")}',
                msg='', files=[str(path)], compress_files=True)
            except:
                print(f"Failed to send compressed file: {e}")
        except:
            print(f"Failed to send: {e}")

    # deleting reserve file
    os.remove(path)

    return None

def int_data(dataframe: pd.DataFrame, fix_on: str) -> pd.DataFrame:
    """
    Forces given data column to ints.\n
    Will NOT work for AWB header.
    """

    assert fix_on != 'AWB', 'Cannot cast AWB header.'

    og_cols = list(dataframe.columns)
    temp_df: pd.DataFrame = dataframe.dropna(axis=0, subset=[fix_on], inplace=False)#drops all awbs with no entry at given col
    temp_df = temp_df.astype({fix_on:  np.int64})#converting all data to int for given col
    drop_cols = list(temp_df.columns)
    print(drop_cols)
    drop_cols.remove(fix_on)
    drop_cols.remove('AWB')
    print(f'cols to drop:\n{drop_cols}') #dropping all cols but awbs, where we are fixing
    temp_df = temp_df.drop(columns=drop_cols)
    #dropping from wfm, then joining
    temp_df = temp_df.set_index('AWB', drop=True, inplace=False)#setting awb as index
    print(f"Data to append:\n{temp_df}")
    dataframe = dataframe.drop(columns=[fix_on]).join(temp_df, on='AWB')#joining with temp data
    #reverting to original order
    dataframe = dataframe.reindex(columns=og_cols)

    return dataframe

if __name__ == "__main__":
    main()