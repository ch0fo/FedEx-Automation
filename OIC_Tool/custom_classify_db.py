"""
This is a custom version of the classify database. Queries the classify database for daily OIC task automation.

Note that you might need to change the database credentials below if they ever get updated.
"""

import oracledb as odb
import pandas as pd
from sqlalchemy import create_engine
import numpy as np
import sys, os
sys.path.append(os.getcwd())
sys.path.append(f"{os.getcwd()}/main_automation_programs")
from main_automation_programs import tools

un = 'classify_query_app'
pw = 'ppa_yreuq_yfissalc'
cs = 'ldap://eusoud.prod.fedex.com/CLASFY_PRD_01_CLASSIFY_S1'

odb.init_oracle_client(lib_dir=r"main_automation_programs\support-files\Oracle\instantclient-basic-windows.x64-23.4.0.24.05\instantclient_23_4")

def execute_query(query: str):
    connection = odb.connect(user=un, password=pw, dsn=cs)
    engine = create_engine('oracle+oracledb://', creator=lambda: connection) #using sqlalchemy for better compatibility
    data: pd.DataFrame = pd.DataFrame()
    try:
        data = pd.read_sql(sql= query, con= engine, parse_dates=['entry_dt'], dtype={'awb_nbr': np.int64, 'billacc': np.int64, 'cad_val': np.float64})
    except ValueError as e:
        print(f"Ecountered error during normal read: {e}")
        #finding faulty lines
        data = pd.read_sql(sql= query, con= engine, parse_dates=['entry_dt'], dtype={'billacc': np.int64, 'cad_val': np.float64})
        data = data.drop(labels = tools.find_OLD(data), inplace=False)

        #fixing datatype
        data = data.astype({'awb_nbr': np.int64})
        
    return data