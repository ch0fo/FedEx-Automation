from typing import List
from tkinter import filedialog
import pandas as pd
from tools import merge_files, get_classify
from tools import get_query
import json, os
from pathlib import Path
import time

"""
Used to find acc nbrs for given awbs. Useful for finding remaining PH accounts shipments.

I use this program to give all awbs an account number, and only keep the ones that PH is responsible for -> find PH work at the end of the month.

Saves output to downloads folder.
"""

def run() -> None:
    #Getting lvs vol files to process
    files: List[str] = list()
    print("Please select files to process")
    for doc in filedialog.askopenfilenames(filetypes= [('CSV Files', '*.csv')]):
        files.append(doc)
    lvs_data: pd.DataFrame = pd.DataFrame()
    lvs_data = merge_files(files, complete=True)
    print(f"Volume data:\n{lvs_data}")

    #Getting classify data
    classify_query: str = get_query(r"main_automation_programs\support-files\queries\classify_COE_Bill-to-acc.sql") #reading saved query
    awbs = list(lvs_data['Tracking Number'])
    classify_data = get_classify(awbs, classify_query) #adding awbs to query, getting data
    print(f"Classify data:\n{classify_data}")

    #Merging data
    complete_df = lvs_data.join(classify_data, on='Tracking Number', lsuffix='_volume', rsuffix='_classify')
    print(f"Merged dataframe:\n{complete_df}")

    #Printing out to downloads
    path = Path(os.path.expanduser("~"))/'Downloads'/f'classify_data_{time.time_ns()}.xlsx'
    complete_df.to_excel(excel_writer=str(path), sheet_name='data', na_rep='null', index=False)
    print(f"Data copied to '{str(path)}'")
    # print(f"Percentage of data found in classify: {(classify_data.shape[0]/lvs_data.shape[0])*100}%")


if __name__ == "__main__":
    run()