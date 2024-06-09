# -*- coding: utf-8 -*-
"""
Created on Fri Jun  7 12:06:56 2024

@author: Hadrien

This script serves to extract each sheet of the source Excel to CSVs.
The goal is to use those CSVs as simple, easy import sources into databases.
"""
import pandas as pd
from datetime import date
import os

#set date to track when files were made
today = str(date.today())

#create a new folder with the current date
newpath = today
if not os.path.exists(newpath):
    os.makedirs(newpath)

#looks for the source Excel in the same folder as the script
read_file = pd.read_excel("PowerBIExcelSource.xlsx", sheet_name=None, header=0)
#defines the location of the result
for sheet_name, data in read_file.items():
    data.to_csv(f"{today}/{sheet_name}.csv", index=False)
    
