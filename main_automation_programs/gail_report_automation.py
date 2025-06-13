from typing import List, Dict
from tkinter import filedialog
import pandas as pd
from pandas import DataFrame
from tools import get_query, find_OLD, merge_files, sql_able_list, get_password, drop_cols
from pathlib import Path
from datetime import datetime, timedelta
import numpy as np
from classify_db import execute_query
from misa_db import execute_query as misa_query
import os
from datetime import datetime

"""
Use this program to automate the Gail Report, which is usually ran once around the 10th, again around the 15th, and one last time before the 23rd.

Feed the two lvs volume files that Gail will send you.

This program will spit out an excel file for each of the following: 

1. One for all FedEx broker shipments that require CAD creation -> Send to John (Scott) Raymond and Cherif Manuel
2. One for all non FedEx broker shipments that require CAD creation -> Send to Gilbert Duplessis
3. One for all GHOST shipments that must be turned into GGG broker -> Send to John (Scott) Raymond and Cherif Manuel
"""

cols_to_keep: List[str] = ['awb_nbr', 'importernme', 'entry_dt', 'billacc', 'cad_val', 'brokr_id', 'duty_bill_to_cd', 'crl_loc', 'coe', 'saccount'] #which cols to keep when creating reduced versions of final sheets

def run(
        ending: str = ((datetime.today().replace(day=1)) - timedelta(days=1)).strftime("'%d-%b-%y'"), #getting last day of last month
        starting: str = (((datetime.today().replace(day=1)) - timedelta(days=1)).replace(day=1)).strftime("'%d-%b-%y'"), #getting first day of last month
        awb_col: str = 'Tracking Number'
    ) -> None:
    """
    The starting date defaults to the first day of last month, and the ending date defaults to the last day of last month, based on the date when you run the program.
    E.g -> If you run this program during May 2025, starting will return '01-Apr-25', and ending returns '30-Apr-25'.
    """

    #Getting lvs vol files to process (Gail's Report lvs files)
    files: List[str] = list()
    print("Please select files to process")
    for doc in filedialog.askopenfilenames(filetypes= [('CSV Files', '*.csv')]):
        files.append(doc)
    lvs_data: pd.DataFrame = pd.DataFrame()
    lvs_data = merge_files(files, dtypes={awb_col: np.int64})
    print("Dropping duplicate awbs from lvs volume")
    lvs_data = lvs_data.drop_duplicates(subset=[awb_col], inplace= False, ignore_index=True) #dropping awb dups
    print(f"Volume data:\n{lvs_data}")

    #Getting classify data
    print("Now fetching Classify data")
    classify_query: str = get_query(r"main_automation_programs\support-files\queries\gail_report_classify.sql") #reading saved query
    classify_data: DataFrame = execute_query(classify_query, ending=ending, starting=starting) #adding awbs to query, getting data
    classify_data = classify_data.drop(find_OLD(classify_data), inplace=False) #dropping OLD awbs
    classify_data = classify_data.astype({'awb_nbr': np.int64}) #converting awbs to int now that OLD was dropped
    print("Dropping duplicate awbs from Classify volume")
    classify_data = classify_data.drop_duplicates(subset=['awb_nbr'], inplace= False, ignore_index=True) #dropping dups
    classify_data = classify_data.set_index(keys='awb_nbr', drop=False) #setting awb as index to prepare for awb drop
    print(f"Classify data:\n{classify_data}")

    #Finding classify data not in Gail's Report
    missing_volume: set = set(classify_data['awb_nbr']) - set(lvs_data[awb_col]) #getting all awbs missing from gail's report
    awbs_to_drop: set = set(classify_data['awb_nbr']) - missing_volume #getting all awbs to drop so that only the missing awbs are kept
    remaining_classify_data: DataFrame = classify_data.drop(awbs_to_drop, inplace= False)
    #Fixing date col
    remaining_classify_data['entry_dt'] = pd.to_datetime(remaining_classify_data['entry_dt'], format="%Y-%m-%d") #reading date in
    remaining_classify_data['entry_dt'] = remaining_classify_data['entry_dt'].dt.strftime("%d-%b-%Y") #formatting date to '05-May-25' format

    print(f"Remaining Classify data after match with Gail's Report (all Classify volume not in Gail's Report):\n{remaining_classify_data}")
    if len(remaining_classify_data.index) == 0:
        print("No missing shipments were found in Gail's Report.")
        return 0 #exit program if there are no awbs left to check for movements

    #Checking movements on remaining shipments
    print("Checking for Canada movements")
    movements_query: str = get_query(r"main_automation_programs\support-files\queries\CA_Delivery.sql").format(
        awbs = sql_able_list(remaining_classify_data['awb_nbr'].to_list(), logic='IN', variable='AL1.shp_trk_nbr')
        ) #formatting query to contain awbs we want to search for
    movements_data: DataFrame = misa_query(movements_query, pw=get_password(), date_query=False).astype({'shp_trk_nbr': np.int64}) #fetching misa, converting awbs to ints
    print(f"Canada movements data: \n {movements_data}") #all theses awbs have movement

    #Checking CRNs
    print("Getting CRNs")
    check_crns: set = set(remaining_classify_data['awb_nbr']) - set(movements_data['shp_trk_nbr']) #getting awbs with no movement, so we can find CRNs for them
    print(f"Checking {len(check_crns)} awbs for CRNs")
    crns_data: DataFrame = DataFrame()
    crns_query: str = get_query(r"main_automation_programs\support-files\queries\Find_CRN.sql")
    if len(check_crns) != 0: #if there are crns to check for movement (not all shipments had movement)
        crns_query = crns_query.format(awbs = sql_able_list(list(check_crns), logic='IN', variable='AL1.AWB_NBR'))
        crns_data = misa_query(crns_query, pw = get_password(), date_query=False)
        crns_data = crns_data.drop_duplicates(['TRACKING_NBR'], inplace=False).astype({'TRACKING_NBR': np.int64}) #dropping crn dups
        crns_data = crns_data.set_index('TRACKING_NBR', drop=False, inplace=False) #setting crn tracking as index
        print(f"Whole CRNs data: \n{crns_data}")
    
    #Checking CRNs movement
    if len(crns_data.index) != 0: #looking for movements when crns were found
        crns_movement_query: str = get_query(r"main_automation_programs\support-files\queries\CA_Delivery.sql").format(
            awbs = sql_able_list(crns_data['TRACKING_NBR'].to_list(), logic='IN', variable='AL1.shp_trk_nbr')
        )
        crns_movement_data: DataFrame = misa_query(crns_movement_query, pw = get_password(), date_query=False).astype({'shp_trk_nbr': np.int64}) #getting crns movements, transforming results to ints
        crns_movement_data = crns_movement_data.drop_duplicates('shp_trk_nbr', inplace=False, ignore_index=True) #dropping dups
        #Now have a list of all CRNs with movement
        print(f"CRNs movement data: \n{crns_movement_data}")
        
        #Finding CRNs with no movement -> Want to remove all crns with NO movement from my CRNS data, so that only the master awbs that have CRNS with movement are kept
        crns_no_movement: set = set(crns_data['TRACKING_NBR']) - set(crns_movement_data['shp_trk_nbr']) #Taking all CRNS awbs, removing CRNS with movement from it, to have a set of all CRN awbs with no movement
        crns_data_movement_only: DataFrame = crns_data.drop(list(crns_no_movement), inplace=False) #now only have awbs with movement
        print(f"CRNs data movement only: \n{crns_data_movement_only}")
        master_awbs_movement: DataFrame = crns_data_movement_only.drop(columns=['TRACKING_NBR'], inplace=False).reset_index(drop=True, inplace=False).rename(columns={'AWB_NBR': 'shp_trk_nbr'}, inplace=False)
        #The line above does the following: drops all CRN awbs with no movement, removes the CRNS column from the dataframe, renames the master awb column from 'AWB_NBR' to 'shp_trk_nbr'
        master_awbs_movement = master_awbs_movement.drop_duplicates('shp_trk_nbr', inplace=False, ignore_index=True).astype({'shp_trk_nbr': np.int64}) #removes any duplicated master awbs, and makes sure they are stored as ints
        print(f"Master Awbs only: \n{master_awbs_movement}")

        #Now, need to append these new master awbs with movement to master movements list
        movements_data = pd.concat([movements_data, master_awbs_movement], ignore_index=True)
        print(f"All awbs with movement: \n {movements_data}")
    
    all_results: Dict[str, DataFrame] = dict() #storing all results here, along with their key (name)

    #Now, need to find GHOST shipments, FedEx shipments missing CAD, and CBI shipments missing CAD
    ghost_shipments: DataFrame = remaining_classify_data.drop(movements_data['shp_trk_nbr'].to_list(), inplace=False) #drop all awbs with movement, left with all true GHOST data
    all_results['ghosts'] = ghost_shipments #storing ghost results

    #Getting FedEx, Non FedEx shipments
    all_movement_shipments: DataFrame = remaining_classify_data.drop(ghost_shipments['awb_nbr'].to_list(), inplace=False) #all shipments with movement

    #Getting FedEx brokers
    fedex_broker_shipments: DataFrame = all_movement_shipments[(all_movement_shipments['brokr_id'] == 'FEC') | (all_movement_shipments['brokr_id'] == 'FON')] #gets all fedex brokers
    all_results['fedex'] = fedex_broker_shipments

    #Getting non fedex brokers
    CBI_broker_shipments: DataFrame = all_movement_shipments.drop(fedex_broker_shipments['awb_nbr'].to_list(), inplace=False)
    all_results['cbi'] = CBI_broker_shipments

    print(f"Ghost shipments: \n {ghost_shipments}")
    print(f"FedEx broker shipments missing CAD: \n {fedex_broker_shipments}")
    print(f"Non FedEx broker shipments missing CAD: \n {CBI_broker_shipments}")

    # #Getting reduced dataframes, as most cols are too bulky to include in email NOTE: No longer doing this, just send them the excel file with the complete info.
    # print("Dropping bulky awbs from results")
    # ghost_shipments_reduced: DataFrame = drop_cols(ghost_shipments, cols_to_keep)
    # fedex_broker_shipments_reduced: DataFrame = drop_cols(fedex_broker_shipments, cols_to_keep)
    # CBI_broker_shipments_reduced: DataFrame = drop_cols(CBI_broker_shipments, cols_to_keep)

    #Pasting results to excel files
    for name, data in all_results.items():
        report_path: Path = Path(os.getcwd())/f'{name}_gails_report_{datetime.today().strftime("%d-%b-%Y")}.xlsx' #getting path to save results to
        print(f"Exporting results to '{report_path.as_posix()}'")
        with pd.ExcelWriter(path= report_path, engine='openpyxl', mode='w') as writer: #writing results out to excel file
            data.to_excel(excel_writer=writer, sheet_name=name, na_rep='', index=False)

if __name__ == "__main__":
    run()