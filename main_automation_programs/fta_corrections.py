"""
Automation for FTA corrections.

The queries are kept in the data/queries/FTA folder.

For each FTA correction, their query is run and fed into the FTA corrections macro.

The results are emailed out.

Run on Mondays, Wednesdays, and Fridays.
"""

from typing import List, Dict
from misa_db import execute_query
import datetime
from datetime import datetime, timedelta
from tools import get_query, get_envvar, send_email, get_password
from pandas import DataFrame
import dotenv, os
from pathlib import Path
import shutil
import xlwings as xw
import time
from numpy import nan

def get_month() -> str:
    """
    Fetches corresponding date for past month.
    Used later on to fetch all other queries.

    Returns a string containg the value for the last month, e.g. '2025-02-01' if ran during March, 2025.
    """

    first_day: datetime =  datetime.today().replace(day=1) #returns datetime object for first day of current month
    first_day_last_month: datetime = (first_day - timedelta(days= 1)).replace(day=1) #getting first day of last month by substracting one day from first_day, and replacing that date to its first day

    return first_day_last_month.strftime("'%Y-%m-%d'") #returning with proper formatting

def get_destination(home_path: str = r"main_automation_programs\reports\audits\fta-corrections") -> Path:
    """
    Returns the filepath to use for the new FTA macro.

    home_path: base path to use for storing; will append appropriate filename for today to given home path.\
    
    Also checks to see if path leading up to filename exists.
    If it doesn't exist, creates path for it.
    """

    home = Path(home_path)

    today = datetime.today()
    file_homepath = home/fr"{today.strftime("%B%Y")}"
    file_path = file_homepath/f"FTA Corrections_{today.strftime("%B%d_%Y")}.xlsm"

    if not os.path.exists(file_homepath):
        os.mkdir(file_homepath)

    return file_path

def fix_headings(data: DataFrame, index: str) -> DataFrame:
    """
    Fixes the names and order of the given dataframe's columns, based on the given index.

    The index is an identifier for the results; '0017_150+', 'ROW_20', etc.

    Drops all unnecessary columns, for each dataframe.

    Adds any necessary, empty, new columns to match fixed cols requirements.
    """

    #Dict showing the correct order for headers
    fixed_cols: Dict[str, List[str]] = {
        'CUSMA_40': [
            'AWB_NBR', 'EMPLOYEE_NBR', 'COUNTRY_CD', 'SHIPMENT_VALUE_FLG',
            'ORDER_IN_COUNCIL_DOC_NBR',
            'TARIFF_ANNEX_CD', 'COMMODITY_LINE_NBR',
            'COMMODITY_DESC', 'HARMONIZED_TARIFF_NBR', 'cad_value_amt',
            'DUTY_VALUE', 'CANADIAN_ENTRY_NBR', 'oga_shipment_flg',
            'DATE_DT', 'REL_DATE', 'DUTY_BILL_TO_ACCT_NBR', 'STATE_CD',
            'CLEARANCE_PORT_CD', 'CASUAL_IMPORTER_CD',
            'TOTAL_VALUE_DUTY_AMT',# used to be -> '(((SALES_TAX_AMT+PROVINCIAL_SALES_TAX_AMT)+SPCL_IMPT_MEAS_ACT_TAX_AMT)+DUTY_AMT)' MIGH NEED TO CHANGE, review calculations of tax,
            'DUTY_AMT', 'PROVINCIAL_SALES_TAX_AMT', 'SALES_TAX_AMT',
            'SPCL_IMPT_MEAS_ACT_TAX_AMT', 'rod_flg', 'COMPANY_NM',
            'CONTACT_NM', # We don't do this one anymore 
            'CITY_NM',
            'STATE_CD', 'STATE_PROVINCE_NM',
            'POSTAL_CD',
            'REFERENCE_NOTES_DESC'
            ],

        '0017_non_CUSMA': [
            'AWB_NBR', 'EMPLOYEE_NBR', 'COUNTRY_CD', 'SHIPMENT_VALUE_FLG',
            'MANUF_ORIGIN_COUNTRY_CD', #this one is different from CUSMA_40; replaced OIC with COM
            'TARIFF_ANNEX_CD', 'COMMODITY_LINE_NBR',
            'COMMODITY_DESC', 'HARMONIZED_TARIFF_NBR', 'cad_value_amt',
            'DUTY_VALUE', 'CANADIAN_ENTRY_NBR', 'oga_shipment_flg',
            'DATE_DT', 'REL_DATE', 'DUTY_BILL_TO_ACCT_NBR', 'STATE_CD',
            'CLEARANCE_PORT_CD', 'CASUAL_IMPORTER_CD',
            'TOTAL_VALUE_DUTY_AMT',
            'DUTY_AMT', 'PROVINCIAL_SALES_TAX_AMT', 'SALES_TAX_AMT',
            'SPCL_IMPT_MEAS_ACT_TAX_AMT', 'rod_flg', 'COMPANY_NM', 'CONTACT_NM',
            'CITY_NM', 'STATE_CD', 'STATE_PROVINCE_NM', 'POSTAL_CD',
            'REFERENCE_NOTES_DESC'
            ],

        '0017_150+': [ #identical to CUSMA_40
            'AWB_NBR', 'EMPLOYEE_NBR', 'COUNTRY_CD', 'SHIPMENT_VALUE_FLG',
            'ORDER_IN_COUNCIL_DOC_NBR','TARIFF_ANNEX_CD', 'COMMODITY_LINE_NBR',
            'COMMODITY_DESC', 'HARMONIZED_TARIFF_NBR', 'cad_value_amt',
            'DUTY_VALUE', 'CANADIAN_ENTRY_NBR', 'oga_shipment_flg',
            'DATE_DT', 'REL_DATE', 'DUTY_BILL_TO_ACCT_NBR', 'STATE_CD',
            'CLEARANCE_PORT_CD', 'CASUAL_IMPORTER_CD',
            'TOTAL_VALUE_DUTY_AMT',
            'DUTY_AMT', 'PROVINCIAL_SALES_TAX_AMT', 'SALES_TAX_AMT',
            'SPCL_IMPT_MEAS_ACT_TAX_AMT', 'rod_flg', 'COMPANY_NM', 'CONTACT_NM',
            'CITY_NM', 'STATE_CD', 'STATE_PROVINCE_NM', 'POSTAL_CD',
            'REFERENCE_NOTES_DESC'
        ],

        'ROW_20': [ #identical to CUSMA_40
            'AWB_NBR', 'EMPLOYEE_NBR', 'COUNTRY_CD', 'SHIPMENT_VALUE_FLG',
            'ORDER_IN_COUNCIL_DOC_NBR','TARIFF_ANNEX_CD', 'COMMODITY_LINE_NBR',
            'COMMODITY_DESC', 'HARMONIZED_TARIFF_NBR', 'cad_value_amt',
            'DUTY_VALUE', 'CANADIAN_ENTRY_NBR', 'oga_shipment_flg',
            'DATE_DT', 'REL_DATE', 'DUTY_BILL_TO_ACCT_NBR', 'STATE_CD',
            'CLEARANCE_PORT_CD', 'CASUAL_IMPORTER_CD',
            'TOTAL_VALUE_DUTY_AMT',
            'DUTY_AMT', 'PROVINCIAL_SALES_TAX_AMT', 'SALES_TAX_AMT',
            'SPCL_IMPT_MEAS_ACT_TAX_AMT', 'rod_flg', 'COMPANY_NM', 'CONTACT_NM',
            'CITY_NM', 'STATE_CD', 'STATE_PROVINCE_NM', 'POSTAL_CD',
            'REFERENCE_NOTES_DESC'
        ],

        'CUKTA_CAS': [ #identical to CUSMA_40
            'AWB_NBR', 'EMPLOYEE_NBR', 'COUNTRY_CD', 'SHIPMENT_VALUE_FLG',
            'ORDER_IN_COUNCIL_DOC_NBR','TARIFF_ANNEX_CD', 'COMMODITY_LINE_NBR',
            'COMMODITY_DESC', 'HARMONIZED_TARIFF_NBR', 'cad_value_amt',
            'DUTY_VALUE', 'CANADIAN_ENTRY_NBR', 'oga_shipment_flg',
            'DATE_DT', 'REL_DATE', 'DUTY_BILL_TO_ACCT_NBR', 'STATE_CD',
            'CLEARANCE_PORT_CD', 'CASUAL_IMPORTER_CD',
            'TOTAL_VALUE_DUTY_AMT',
            'DUTY_AMT', 'PROVINCIAL_SALES_TAX_AMT', 'SALES_TAX_AMT',
            'SPCL_IMPT_MEAS_ACT_TAX_AMT', 'rod_flg', 'COMPANY_NM', 'CONTACT_NM',
            'CITY_NM', 'STATE_CD', 'STATE_PROVINCE_NM', 'POSTAL_CD',
            'REFERENCE_NOTES_DESC'
        ],

        'CUKTA_NR': [ #No reference on header, but I am assuming it should be identical to CUSMA_40
            'AWB_NBR', 'EMPLOYEE_NBR', 'COUNTRY_CD', 'SHIPMENT_VALUE_FLG',
            'ORDER_IN_COUNCIL_DOC_NBR','TARIFF_ANNEX_CD', 'COMMODITY_LINE_NBR',
            'COMMODITY_DESC', 'HARMONIZED_TARIFF_NBR', 'cad_value_amt',
            'DUTY_VALUE', 'CANADIAN_ENTRY_NBR', 'oga_shipment_flg',
            'DATE_DT', 'REL_DATE', 'DUTY_BILL_TO_ACCT_NBR', 'STATE_CD',
            'CLEARANCE_PORT_CD', 'CASUAL_IMPORTER_CD',
            'TOTAL_VALUE_DUTY_AMT',
            'DUTY_AMT', 'PROVINCIAL_SALES_TAX_AMT', 'SALES_TAX_AMT',
            'SPCL_IMPT_MEAS_ACT_TAX_AMT', 'rod_flg', 'COMPANY_NM', 'CONTACT_NM',
            'CITY_NM', 'STATE_CD', 'STATE_PROVINCE_NM', 'POSTAL_CD',
            'REFERENCE_NOTES_DESC'
        ],

        'CETA_CAS': [
            'AWB_NBR', 'COMMODITY_LINE_NBR', 'EMPLOYEE_NBR', 'COUNTRY_CD', 'MANUF_ORIGIN_COUNTRY_CD', 'TARIFF_TREATMENT_CD',
            'SHIPMENT_VALUE_FLG', 'COMMODITY_DESC', 'HARMONIZED_TARIFF_NBR', 'cad_value_amt',
            'DUTY_VALUE', 'CANADIAN_ENTRY_NBR', 'oga_shipment_flg',
            'DATE_DT', 'REL_DATE', 'DUTY_BILL_TO_ACCT_NBR', 'STATE_CD',
            'CLEARANCE_PORT_CD', 'CASUAL_IMPORTER_CD',
            'TOTAL_VALUE_DUTY_AMT',
            'DUTY_AMT', 'PROVINCIAL_SALES_TAX_AMT', 'SALES_TAX_AMT',
            'SPCL_IMPT_MEAS_ACT_TAX_AMT', 'rod_flg', 'COMPANY_NM', 'CONTACT_NM',
            'CITY_NM', 'STATE_CD', 'STATE_PROVINCE_NM', 'POSTAL_CD',
            'REFERENCE_NOTES_DESC', 'DUTY_BILL_TO_ACCT_NBR'
        ],

        'CETA_NR': [ #identical to CETA_CAS
            'AWB_NBR', 'COMMODITY_LINE_NBR', 'EMPLOYEE_NBR', 'COUNTRY_CD', 'MANUF_ORIGIN_COUNTRY_CD', 'TARIFF_TREATMENT_CD',
            'SHIPMENT_VALUE_FLG', 'COMMODITY_DESC', 'HARMONIZED_TARIFF_NBR', 'cad_value_amt',
            'DUTY_VALUE', 'CANADIAN_ENTRY_NBR', 'oga_shipment_flg',
            'DATE_DT', 'REL_DATE', 'DUTY_BILL_TO_ACCT_NBR', 'STATE_CD',
            'CLEARANCE_PORT_CD', 'CASUAL_IMPORTER_CD',
            'TOTAL_VALUE_DUTY_AMT',
            'DUTY_AMT', 'PROVINCIAL_SALES_TAX_AMT', 'SALES_TAX_AMT',
            'SPCL_IMPT_MEAS_ACT_TAX_AMT', 'rod_flg', 'COMPANY_NM', 'CONTACT_NM',
            'CITY_NM', 'STATE_CD', 'STATE_PROVINCE_NM', 'POSTAL_CD',
            'REFERENCE_NOTES_DESC', 'DUTY_BILL_TO_ACCT_NBR'
        ],

        'CPTPP_CAS': [ #identical to CETA_CAS
            'AWB_NBR', 'COMMODITY_LINE_NBR', 'EMPLOYEE_NBR', 'COUNTRY_CD', 'MANUF_ORIGIN_COUNTRY_CD', 'TARIFF_TREATMENT_CD',
            'SHIPMENT_VALUE_FLG', 'COMMODITY_DESC', 'HARMONIZED_TARIFF_NBR', 'cad_value_amt',
            'DUTY_VALUE', 'CANADIAN_ENTRY_NBR', 'oga_shipment_flg',
            'DATE_DT', 'REL_DATE', 'DUTY_BILL_TO_ACCT_NBR', 'STATE_CD',
            'CLEARANCE_PORT_CD', 'CASUAL_IMPORTER_CD',
            'TOTAL_VALUE_DUTY_AMT',
            'DUTY_AMT', 'PROVINCIAL_SALES_TAX_AMT', 'SALES_TAX_AMT',
            'SPCL_IMPT_MEAS_ACT_TAX_AMT', 'rod_flg', 'COMPANY_NM', 'CONTACT_NM',
            'CITY_NM', 'STATE_CD', 'STATE_PROVINCE_NM', 'POSTAL_CD',
            'REFERENCE_NOTES_DESC', 'DUTY_BILL_TO_ACCT_NBR'
        ]
    }

    #Adding any applicable missing columns (adds with all empty numpy NaN values.)
    new_data: DataFrame = data.copy(deep=True)
    for header in fixed_cols[index]:
        if header not in data.columns: #checking if header is missing from existing data
            data[header] = nan #for headers missing, create new col with empty values
    
    new_data = data[fixed_cols[index]]
    return new_data #only return cols in the given list, in the given order; gets list from fixed_cols dictionary

def run_corrections() -> None:
    """
    Main function for getting the corrections, running through FTA macro, and emailing results.
    """

    start: str = get_month() #getting month to use for queries. For example, if current month is November 2024, will get '2024-10-01' -> Fetches everything for all of last month
    end: str = (datetime.today() - timedelta(days=1)).strftime("'%Y-%m-%d'") #day to use as upper limit of query (day before the queries are ran)
    queries_path = Path(r"main_automation_programs\support-files\queries\FTA") #general path for FTA queries
    results: Dict[str, DataFrame] = dict() #keep results in this directory; maps a query name to its results

    for query in os.listdir(queries_path): #getting results for each query

        #Get results for current query.
        name: str = query.split('.')[0]
        print(f"Querying for '{name}'")
        results[name] = execute_query(
                                        query=get_query(str(queries_path/query)), #getting query path for current query
                                        pw = get_envvar('pw'),
                                        starting= start,
                                        ending= end
                                    )
        
        print(f"Query for '{name}' fetched successfully.\n")
        print(results[name])

    #fixing order of cols in results
    for index, dataframe in results.items():
        results[index] = fix_headings(dataframe, index)
    
    # creating copy of last fta macro, and saving as new FTA path
    new_corrections_path = Path(shutil.copyfile(src=get_envvar('last_fta_filepath', envfile_path=dotenv.find_dotenv()), #getting last fta path
                                           dst=get_destination())) #creating new file with today's formatting
    
    #Pasting results, running macros
    empty_fta: bool = True
    with xw.App(visible=False) as app: #initializing xlwing's excel app object
        fta_wb = app.books.open(fullname= new_corrections_path) #opening macro path

        #getting macros
        reset = fta_wb.macro("Module1.resetShts")
        run_fta = fta_wb.macro("Module1.Button1_Click")

        #resetting sheets
        print("Resetting macro")
        reset()

        #pasting results into sheets
        print("Pasting results")
        for sheet in fta_wb.sheet_names:
            if sheet not in results.keys() or results[sheet].shape[0] == 0: #skipping when sheetname is not in results, or when results are empty (empty dataframe)
                continue

            curr_sheet: xw.Sheet = fta_wb.sheets[sheet] #creating new sheet object for current sheet
            values = results[sheet].values.tolist()
            curr_sheet.range('D2').value =  values #pasting values into sheet

        #Running main macro
        fta_wb.save()
        print("Running main macro")
        run_fta()

        #Cleaning results, so there are only 150 rows per results, and checking for empty results
        print("Cleaning results")
        for sheet in fta_wb.sheet_names:
            if sheet not in results.keys() and sheet != 'CUSMA_40_150':
                print(f"Skipping sheet '{sheet}'")
                continue #skip sheets not related to results

            curr_sheet: xw.Sheet = fta_wb.sheets[sheet] #creating new sheet objet

            #checking for empty sheet
            curr_val: object = curr_sheet['D2'].value

            if curr_val is None:
                print(f"None value for sheet '{sheet}'")
            else:
                print(f"Value found for sheet '{sheet}': {curr_val}, {type(curr_val)}")
                empty_fta = False #When a value is found for any sheets, corrections are emailed out

            #Cleaning sheet (leaving only top 150 corrections per sheet)
            clear_range: xw.Range = curr_sheet.range('D152:AJ152').expand('down') #selecting range to clear
            clear_range.clear_contents() #clearing range
            print(f"Cleaned sheet '{sheet}'")

        #saving and closing
        print("Saving, closing")
        fta_wb.save()
        fta_wb.close
    #Emailing results
    try:
        if not empty_fta: 
            send_email(f'FTA_CORRECTIONS_0003 {datetime.today().strftime("%d-%b-%Y")}', files=[str(new_corrections_path)])
            print("Valid corrections")
        else: print("Unable to email corrections; empty corrections file.")
    except Exception as e:
        print(f"Failed to email corrections, see details: \n \n {e} \n \n")
    else:
        #updating path for last used fta_macro
        dotenv.set_key(dotenv_path=dotenv.find_dotenv(), key_to_set='last_fta_filepath', value_to_set= str(new_corrections_path))
        #Note that the path will only update if the corrections are emailed successfully.
        
    return None

if __name__ == "__main__":
    run_corrections()