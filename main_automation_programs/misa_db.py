"""
Used for queying the MISA database from python.

Please note that you might need to change the credentials below if they ever get updated (you will probably receive an email warning you that the credentials are changing)
"""

import oracledb as odb
import pandas as pd
from sqlalchemy import create_engine
import time
from tools import get_envvar
import teradatasql

un = get_envvar('misa-username')
cs = 'edwmiscop1.prod.fedex.com'

odb.init_oracle_client(lib_dir=r"main_automation_programs\support-files\Oracle\instantclient-basic-windows.x64-23.4.0.24.05\instantclient_23_4")

def execute_query(query: str, pw: str, dates: str = '', starting: str = '', ending: str = ''):
    """
    Executes given query in MISA DB.

    Note that you may pass a single date, or a starting and ending, but not both.

    Defaults to starting and ending being empty.
    """
    t = time.time()
    print(f"Fetching '{cs}' with query '{query[:min(len(query), 200)].replace('\n', ' ')}' (complete query may not be shown)")
    #fetching
    with teradatasql.connect(host=cs, user=un, password=pw) as connection:
        engine = create_engine('teradatasql://', creator=lambda: connection) #using sqlalchemy for better compatibility
        if starting != '' and ending != '':
            print(f"Fetching for dates {starting} to {ending}")
            curr_qry: str = query.format(starting = starting, ending = ending)
        else:
            print(f"Fetching for dates {dates}")
            curr_qry: str = query.format(dates = dates)
        data = pd.read_sql_query(sql=curr_qry, con=engine)
    print(f"Time taken to run MISA query: {time.time()-t} secs")
    return data