"""
Used for queying the MISA database from python.

Please note that you might need to change the credentials below if they ever get updated (you will probably receive an email warning you that the credentials are changing)
"""

import oracledb as odb
import pandas as pd
from sqlalchemy import create_engine
import time
import sys, os
sys.path.append(os.getcwd())
sys.path.append(f"{os.getcwd()}/main_automation_programs")
import teradatasql

def execute_query(query: str, pw: str, dates: str = '', starting: str = '', ending: str = '', date_query: bool = True):
    """
    Executes given query in MISA DB.

    Note that you may pass a single date, or a starting and ending, but not both.

    Defaults to starting and ending being empty.
    """
    #Initializing oracledb
    from tools import get_envvar
    odb.init_oracle_client(lib_dir=get_envvar('oracle_install_path'))

    t = time.time()
    un = get_envvar('misa-username')
    cs = 'edwmiscop1.prod.fedex.com'
    print(f"Fetching '{cs}' with query '{query[:min(len(query), 200)].replace('\n', ' ')}' (complete query may not be shown)")

    #fetching
    with teradatasql.connect(host=cs, user=un, password=pw) as connection:
        engine = create_engine('teradatasql://', creator=lambda: connection) #using sqlalchemy for better compatibility
        if date_query and starting != '' and ending != '':
            print(f"Fetching for dates {starting} to {ending}")
            curr_qry: str = query.format(starting = starting, ending = ending)
        elif date_query:
            print(f"Fetching for dates {dates}")
            curr_qry: str = query.format(dates = dates)
        else:
            print(f"Fetching MISA")
            curr_qry: str = query
        data = pd.read_sql_query(sql=curr_qry, con=engine)
    print(f"Time taken to run MISA query: {time.time()-t} secs")
    return data