"""
Used for querying the Classify database from python.

Please note that you might need to change the credentials below if they ever get updated (you will probably receive an email warning you that the credentials are changing)
"""

import oracledb as odb
import pandas as pd
from sqlalchemy import create_engine
import numpy as np
import time

un = 'classify_query_app'
pw = 'ppa_yreuq_yfissalc'
cs = 'ldap://eusoud.prod.fedex.com/CLASFY_PRD_01_CLASSIFY_S1'

#initialize the oracle client, needed to run the queries
odb.init_oracle_client(lib_dir=r"main_automation_programs\support-files\Oracle\instantclient-basic-windows.x64-23.4.0.24.05\instantclient_23_4")

def execute_query(query: str, starting: str = '', ending: str = ''):
    t = time.time()
    print(f"Fetching '{cs}' with query '{query[:min(len(query), 200)].replace('\n', ' ')}' (complete query may not be shown)")
    connection = odb.connect(user=un, password=pw, dsn=cs)
    engine = create_engine('oracle+oracledb://', creator=lambda: connection) #using sqlalchemy for better compatibility
    #Checking dates
    if starting != '' and ending != '':
        print(f"Fetching for dates {starting} to {ending}")
        query = query.format(starting = starting, ending = ending)
    try:
        data = pd.read_sql(sql= query, con= engine, dtype={'awb_nbr': np.int64, 'bill_to_acc': np.int64})
    except:
        data = pd.read_sql(sql=query, con=engine)
    print(f"Time taken to run Classify query: {time.time()-t} secs\n{data}")
    return data