"""
Small tool to split a given .txt file containing awbs (one awb per line) into chunks of 35k awbs (max for processing awbs in Hyperion queries)
Useful when refreshing reports -> Used for refreshing the BAS Cases / Surtax report.
"""

import os.path
from datetime import date
import csv

_BILLS_PER_CHUNK_ = 35000

def main():
    path = fetch_path()
    with open(path, newline='') as f:
        reader = csv.reader(f)
        airway_bills = list(reader)
    _NBILLS_ = len(airway_bills)
    #creating chunks
    chunks = []
    for i in range(0, _NBILLS_, _BILLS_PER_CHUNK_):
        chunks.append(airway_bills[i:i+_BILLS_PER_CHUNK_])
    #creating txt
    create_chunks(chunks)

def fetch_path():
    while True:
        path = input("Airwaybills path: ")
        if "\"" in path:
            path = path.replace("\"", "")
        if os.path.isfile(path):
            return path
        print("\nFile not found, try again.")

def create_chunks(chunks):
    paths = []
    tday = date.today()
    chunks_path = "chunks_{0}".format(tday)
    movements_path = r"main_automation_programs\awb_splits"
    path = "{0}\\{1}".format(movements_path, chunks_path)
    if not os.path.exists(path):
        os.makedirs(path)
    
    i = 1
    for chunk in chunks:
        filename = "chunk_{0}_{1}.csv".format(i, tday)
        file_path = "{0}\\{1}".format(path, filename)
        with open(file_path, 'w', newline='') as myfile:
            wr = csv.writer(myfile)
            wr.writerows(chunk)

        print("created chunk at '{0}'".format(file_path))
        paths.append(file_path)
        i += 1
    
    for file in paths:
        os.rename(file, file.replace(".csv", ".txt"))

if __name__ == "__main__":
    main()