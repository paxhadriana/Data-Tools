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

#Specify the name of the source file here
SOURCE_FILE = "PowerBIExcelSource.xlsx"

#set date to track when files were made
today = str(date.today())

#create a new folder with the current date if it doesnt exist yet
newpath = today
if not os.path.exists(newpath):
    os.makedirs(newpath)

#reads for the source Excel
read_file = pd.read_excel(SOURCE_FILE, sheet_name=None, header=0)
#creates .csvs from the source file sheets and adds them to the folder
for sheet_name, data in read_file.items():
    data.to_csv(f"{today}/{sheet_name}.csv", index=False)

    
