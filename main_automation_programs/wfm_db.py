"""
Used for queying the WFM database from python (CTA AIR database).

Please note that you might need to change the credentials below if they ever get updated (you will probably receive an email warning you that the credentials are changing)
"""

import oracledb as odb
import pandas as pd
from sqlalchemy import create_engine
import numpy as np
import time
from typing import List
import sys, os
sys.path.append(os.getcwd())
sys.path.append(f"{os.getcwd()}/main_automation_programs")

un = 'CTA_AIR_RO_APP'
pw = 'pQMh1Kx_AaXeMtKbvbmzUUUYkB_b87'
cs = 'CTA_PROD_FXE_CAN_DBA_SVC.prod.iaas.fedex.com'
hostname = 'P100375-scan.prod.iaas.fedex.com'
port = 1526

def execute_query(query: str) -> pd.DataFrame:

    #Initializing oracledb
    from tools import get_envvar
    odb.init_oracle_client(lib_dir=get_envvar('oracle_install_path'))

    t = time.time()
    params = odb.ConnectParams(host=hostname, port=port, service_name=cs)
    con = odb.connect(user=un, password=pw, params=params)
    print(f"Db version: {con.version}")
    engine = create_engine('oracle+oracledb://', creator=lambda: con) #using sqlalchemy for better compatibility
    print(f"Fetching '{cs}' with query '{query[:min(len(query), 200)]}' (complete query may not be shown)")
    try: #reading with ideal reserve settings
        data = pd.read_sql(sql=query, con=engine, parse_dates=['entry_dt', 'crt_dt', 'assign_dt','updt_dt'],
                           dtype={'awb': np.int64})
    except:
        print(f'Unable to read dataframe with settings, reading basic instead')
        data = pd.read_sql(sql=query, con=engine)
    print(f'Time taken to fetch query: {time.time()-t} secs')
    return data