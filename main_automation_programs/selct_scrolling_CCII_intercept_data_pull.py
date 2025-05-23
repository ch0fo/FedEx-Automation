"""
This program is used to fetch CCII database entries -> Used for the SELCT Reclassifying productivity report.
Fetches CCII entries in the given date range and returns a .csv file.

Use this program to refresh the SELCT Reclass productivity report.
"""

#Libraries
import teradatasql
import pandas as pd
import datetime
from sqlalchemy import create_engine
from typing import List, Literal
from sqlalchemy.exc import OperationalError
import numpy as np
import time
from tools import get_query, get_envvar
from csv import QUOTE_STRINGS

_MODES = Literal['csv', 'excel']

def get_intercept_data(
    query: str,
    password: str,
    custom_last_day: datetime.datetime = datetime.datetime(1900, 1, 1),
    first_day: datetime.datetime = datetime.datetime(2024,8,30),
    days: int = 31,
    day_shift: int = 0,
    query_host: str = "edwmiscop1.prod.fedex.com",
    username: str = get_envvar('misa-username'),
    save_dataframe_as: _MODES = 'csv',
    return_dataframe: bool = False,
    spooling_delay: int = 5
    ) -> pd.DataFrame | None:
    """
    Gets Intercept data from MISA database.
    Please note that specifying a custom last day overwrites the 'days' argument.
    Please provide 'last_day' as a datetime object. E.g: if you want to pass August 30th, 2024 as the last day, provide datetime(2024, 8, 30).
    The 'day_shift' argument is used to specify how many days to include per round of query fetching. This is necessary because you may run out of spool space
    when running the queries. You can play around with it to maximize your days per query, so that it takes less time to get all the data.

    The 'spooling_delay' variable can be used to wait out the executions, so that time is given for spool to clear.
    The 'save_dataframe_as' variable can be used to specify if you want to save the resulting dataframe as an excel file or a csv file.
    Please note that 'return_dataframe' overrides 'save_dataframe_as'. If return dataframe is true, dataframes are not saved.
    """

    #Settings for query
    last_day: datetime.datetime = first_day + datetime.timedelta(days=days) #Getting last date to process
    if custom_last_day != datetime.datetime(1900,1,1): #checking if a custom last day was specified (overwrites 'days' argument)
        last_day = custom_last_day

    #Checking dates integrity
    assert first_day <= last_day, "Your last day cannot happen before your first day"

    #Getting query data
    dataframe: pd.DataFrame = pd.DataFrame()

    start: datetime.datetime = first_day #initialize first interval to begin with the first day
    end: datetime.datetime = datetime.datetime(1900, 1, 1) #initializing last day for interval
    successful_fetch: bool = False #defining initial fetch condition
    while not (successful_fetch and end == last_day): #keep executing until we are at the last day and we had a successful fetch (with end as the last day)
        if start > last_day: #break out of loop when start date is over last day. This can only happen when we were unable to fetch the last day, in which case we need to exit the loop
            break
        successful_fetch = False #resetting successful fetch checker
        curr_shift: int = min(day_shift, (last_day - start).days) #getting either day shift as passed, or difference between current start and last day to process, whichever is smaller
        print(f"Using day shift: {curr_shift}")
        while not successful_fetch: #keep executing until successful fetch (main reasons to fail are spooling error; out of spooling space)
            end = start + datetime.timedelta(days=curr_shift)
            end = min(end, last_day) #checking that we are not going over last day

            #Defining dates to use
            starting: str = f"'{start.strftime("%Y-%m-%d")}'"
            ending: str = f"'{end.strftime("%Y-%m-%d")}'"
            try:
                with teradatasql.connect(host=query_host, user=username, password=password) as connection:
                    engine = create_engine('teradatasql://', creator=lambda: connection) #using sqlalchemy for better compatibility
                    print(f"Fetching dates {starting} to {ending}")
                    curr_qry: str = query.format(start_date = starting, end_date = ending)
                    print(f'Current query: {curr_qry[:min(300, len(curr_qry))].replace('\n', ' ')}')
                    data = pd.read_sql_query(sql=curr_qry, con=engine,
                                                dtype={
                                                    'AWB_NBR': np.int64,
                                                    'EMPLOYEE_NBR': np.int64
                                                }
                                            )
            except OperationalError as e:
                print(f"Unable to fetch with given day shift (most likely ran out of spooling space)\nReducing day shift by half and refetching. {type(e)}")
                # print(e)
                if curr_shift == 0: #when unable to fetch again with shift already at zero, skip current day. Also skip if time difference is zero (might alredy be doing feb 14 - feb 14, for example, but time shift is over 0)
                    print(f"Skipping day {starting}, unable to fetch.")
                    start += datetime.timedelta(days=1) #artificially adding one day to start date, so the day is skipped
                    # start = min(start, last_day) #checking we are not going over last day
                    break #breaking out of inner while loop to reset day shift
                curr_shift = int(curr_shift/2) #reducing day shift by half
            except Exception as e:
                print(f"Unexpected exception, {type(e)}, {e}")
            else:
                successful_fetch = True

            #Modifying start date after successful fetch
            if successful_fetch:
                print(f"Successful fetch.")
                start += datetime.timedelta(days=1+curr_shift) #shifting start by one day, plus the successful day shift
                start = min(start, last_day) #checking we are not going over last day

            #Wait for spool to clear, when either unsuccessful fetch or last date has not been reached
            if not successful_fetch or end != last_day:
                print("Waiting for spool to clear")
                curr_clear: int = spooling_delay
                while curr_clear > 0: #waiting for spool to clear
                    print(curr_clear)
                    time.sleep(1)
                    curr_clear -= 1       

        #Appending data when successfully fetched any data
        if successful_fetch:
            print(f"Rows fetched: {data.shape[0]} for date range {starting} - {ending}")
            if dataframe.empty:
                dataframe = data
            else:
                dataframe = pd.concat([dataframe, data], ignore_index=True) #If dataframe is not empty, concatenate to existing one

    #Returning or saving to excel files
    if return_dataframe:
        return dataframe

    else:
        if save_dataframe_as == 'excel':
            print("Saving dataframe(s) to excel file(s).")
            dataframes: List[pd.DataFrame] = split_dataframe(dataframe) #getting list of dataframes to save

            curr_dataframe: int = 1
            for df in dataframes: #saves each dataframe into its own excel file
                with pd.ExcelWriter(f'intercept_{first_day.strftime("%d_%b_%Y")}-{last_day.strftime("%d_%b_%Y")}_{curr_dataframe}.xlsx', mode='w') as writer:
                    df.to_excel(excel_writer=writer, sheet_name='intercept', na_rep='null', index=False)
                    print(f"Saved {curr_dataframe} of {len(dataframes)} dataframes.")
                curr_dataframe += 1
            
            return None

        elif save_dataframe_as == 'csv':
            print("Saving dataframe to csv file.")
            with open(f'intercept_{first_day.strftime("%d_%b_%Y")}-{last_day.strftime("%d_%b_%Y")}.csv', mode='w', newline="") as file:
                dataframe.to_csv(path_or_buf=file, na_rep='null', index=False, quoting=QUOTE_STRINGS)
            # with pd.ExcelWriter(f'intercept_{first_day.strftime("%d_%b_%Y")}-{last_day.strftime("%d_%b_%Y")}_{curr_dataframe}.xlsx', mode='w') as writer:
            #         df.to_excel(excel_writer=writer, sheet_name='intercept', na_rep='null', index=False)
            #         print(f"Saved {curr_dataframe} of {len(dataframes)} dataframes.")

            return None
    
def split_dataframe(source_dataframe: pd.DataFrame, split_length: int = 1048576) -> List[pd.DataFrame]:
    """
    Excel is limited to 1,048,576 rows per sheet. This function splits given dataframe into as many dataframes of row length 1,048,576 that can be created,
    or splits with whatver length was specified. Returns list of all split dataframes.
    """

    dataframes: List[pd.DataFrame] = list()

    for index in range(0, len(source_dataframe), split_length):
        dataframes.append(source_dataframe.iloc[index:index+split_length])

    return dataframes

if __name__ == '__main__':
    get_intercept_data(query= get_query(query_path=r'main_automation_programs\support-files\queries\intercept_data.sql'),
                       #Modify the dates below to fetch for the time window you want
                       first_day= datetime.datetime(2025,5,10), custom_last_day=datetime.datetime(2025,5,20), day_shift=4,
                       password=get_envvar(var_name='pw'))