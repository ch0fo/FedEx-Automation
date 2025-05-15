"""
Runs audits found in data/queries/audits

Note that it does not run compliant audits, according to the 'compliance' dictionary below.

Non-compliant audits are run in MISA, their results saved into excel files, and emailed out individually.

Emailed audits are processed using power automate and uploaded into the GTS Audit Team sharepoint.
"""

from typing import List, Dict
from datetime import datetime, timedelta
from pathlib import Path
import os
from pandas import DataFrame
from misa_db import execute_query
from tools import get_query, get_password, send_email
import numpy as np
import shutil
import xlwings as xw
import dotenv
from numpy import nan

compliant: Dict[str, bool] = {
    '2a': True,
    '2b': True,
    '2c': False,
    '3a': True,
    '3b': True,
    '7': False
}

complete_name: Dict[str, str] = {
    '2a': "2a. USMCA LVS ROD $40_$150",
    '2b': "2b. USMCA LVS ROD over $150",
    '2c': "2c. CETA",
    '3a': "3a. USMCA Weekly Com LVS $40 to $150",
    '3b': "3b. USMCA Weekly Com LVS over $150",
    '7': "7. CUKTCA_Commercial"
}

def fix_headers(data: DataFrame, index: str) -> DataFrame:
    """
    Fixes the names and order of the given dataframe's columns, based on the given index (audit name).

    The index is an identifier for the results; '2c', '7', etc.

    Drops all unnecessary columns, for each dataframe.

    Adds any necessary, empty, new columns to match fixed cols requirements.
    """

    #Dict showing the correct order for headers
    fixed_cols: Dict[str, List[str]] = {
        '2c': [
            'Awb_Nbr',
            'N_of_Lines',
            'Value_Flg',
            'Tariff_Cd',
            'COO',
            'COE',
            'Importer_Nm',
            'Customs_Value_Amt',
            'Currency_Cd',
            'Duty_Bill_To_Acct_Nbr',
            'Clr_Loc',
            'Entry_Dt',
            'Release_Dt',
            'Rod_Flg',
            'Employee',
            'Duty_Amt',
            'Status',
            'Tariff_Annex_Cd',
            'PST',
            'GST',
            'Exchange_Rate',
            'Customs_Value_In_CAD'
            ],

        '7': [
            'Awb_Nbr',
            'N_of_Lines',
            'Duty_Bill_To_Acct_Nbr',
            'Value_Flg',
            # 'Tariff_Cd',
            'COE',
            'COO',
            'Importer_Nm',
            'Canadian_Entry_Nbr',
            'Business_Nbr',
            # 'Customs_Value_Amt',
            # 'Currency_Cd',
            # 'Clr_Loc',
            'Entry_Dt',
            'Release_Dt',
            'Rod_Flg',
            'Customs_Value_In_CAD',
            'Employee',
            'Tariff_Annex_Cd',
            # 'Status', 
            'PST',
            'GST',
            'CustomsAmt_BKR_CAD',
            'OIC',
            'SIMA'
            # 'Exchange_Rate',
            ]
    }

    #Adding any applicable missing columns (adds with all empty numpy NaN values.)
    new_data: DataFrame = data.copy(deep=True)
    for header in fixed_cols[index]:
        if header not in data.columns: #checking if header is missing from existing data
            data[header] = nan
    
    new_data = data[fixed_cols[index]]
    return new_data #only return cols in the given list, in the given order; gets list from fixed_cols dictionary

def run_audits(audits_query_path: Path = Path(r"main_automation_programs\support-files\queries\quality_audits"),
               general_audits_path: Path = Path(r"main_automation_programs\reports\audits\quality-audits")) -> None:
    """
    Runs all non-compliant audits.

    The audits_query_path is the path to use to look for the audits sql queries; specify your own custom path.
    The general_audits_path is used to store all audit results, and to get the audit templates.

    Note that the path general_audits_path/templates is expected to contain all audit templates (empty templates that the program makes a copy of)
    """

    today = datetime.today() #getting today's date

    month = today.strftime("%B%Y") #saving today's month and year
    ending = today.strftime("'%Y-%m-%d'") #ending date for queries
    starting = (today - timedelta(days=4)).strftime("'%Y-%m-%d'") #substracting four days from end date to get starting date

    #creating path for current month, in case it does not exist yet
    if not os.path.exists(general_audits_path/month):
        os.mkdir(general_audits_path/month)

    for query_file in os.listdir(audits_query_path): #iterating filenames in the given queries' path
        audit_name = query_file.split('.')[0] #getting query file name
        
        if compliant[audit_name]: #checking if the current audit is not compliant (only runs non-compliant audits)
            continue #when compliant, skips the current query
        print(f"Now reading non compliant '{audit_name}'")

        dotenv.load_dotenv(dotenv_path=dotenv.find_dotenv(), override=True) #loads closest .env file
        if audit_name == '3a' and os.environ['3a_skip']:
            dotenv.set_key(dotenv_path=dotenv.find_dotenv(), key_to_set='3a_skip', value_to_set='')
            print(f"Skipping '{audit_name}'")
            continue #skipping 3a for current week, since it is only run once every two weeks

        elif audit_name == '3a':
            dotenv.set_key(dotenv_path=dotenv.find_dotenv(), key_to_set='3a_skip', value_to_set='1') #modifying env var so that 3a is skipped next week
            print(f"Setting '{audit_name}' to true")

        #running query for current audit
        audit_results: DataFrame = execute_query(query= get_query((audits_query_path/query_file).as_posix()),
                                                 pw= get_password(),
                                                 starting=starting,
                                                 ending=ending
                                                )
        print(f"Query for '{audit_name}' fetched successfully.\n")
        audit_results = audit_results[:150] #keeping first 150 rows only
        print(audit_results)

        #skipping if empty results
        if audit_results.shape[0] == 0:
            print(f"Skipping for audit '{audit_name}' since results are empty\n")
            continue
        
        #Fixing header order / adding extra cols (even if empty) to match template format
        audit_results = fix_headers(audit_results, audit_name)

        #looking for correct template
        curr_template = ''
        for template_name in os.listdir(general_audits_path/'templates'): #looking for current template
            if audit_name in template_name:
                curr_template = template_name

        #creating copy of correct template
        audit_path = Path(shutil.copyfile(src=general_audits_path/'templates'/curr_template, #getting correct template
                                            dst=general_audits_path/month/curr_template.replace('.xlsm', f'{today.strftime("%b_%d_%Y")}.xlsm'))) #creating new file with today's date

        #pasting query results
        with xw.App(visible=False) as app:
            audit_wb = app.books.open(fullname=audit_path) #opening newly created file

            #pasting results into sheets
            print("Pasting results")
            sheet_name = ''
            for sheet in audit_wb.sheet_names:
                if audit_name in sheet:
                    sheet_name = sheet #finding correct sheet name to paste in

            curr_sheet: xw.Sheet = audit_wb.sheets[sheet_name] #creating new sheet object for current sheet
            values = audit_results.values.tolist()
            curr_sheet.range('A2').value =  values #pasting values into sheet

            #saving and closing
            print("Saving, closing")
            audit_wb.save()
            audit_wb.close()

        #email results
        print("Emailing results")
        send_email(f"AUDITS_AUTOMATION_{audit_name} {today.strftime("%b-%d-%Y")}", files=[audit_path],
                   msg=complete_name[audit_name])

if __name__ == '__main__':
    run_audits()