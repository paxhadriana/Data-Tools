# -*- coding: utf-8 -*-
"""
Created on Fri Jun  7 12:06:56 2024

@author: Hadrien

This script serves to extract each sheet of a source Excel file to CSVs.
The goal is to use this as one step in a workflow where data is prepared in Excel then exported into a database via CSVs. 
"""
import pandas as pd
from datetime import date
import os

#Specify the name of the default source file here
DEFAULT = "PowerBIExcelSource.xlsx"

#set date to track when files were made
TODAY = str(date.today())

def main():
    folder()
    extract()

#then create a new folder with the current date if it doesnt exist yet
def folder():
    newpath = TODAY
    if not os.path.exists(newpath):
        os.makedirs(newpath)

#verify the source then extract it to the folder
def extract():
    print("The file to be extracted is currently set to: " + DEFAULT)
    answer = input("Is this correct? (Y/N) ")
    if str(answer) == "N":
        source_file = input("Enter the name of the source file including the .extension: ")
    elif str(answer) == "Y":
        source_file = DEFAULT
#reads for the source Excel
    read_file = pd.read_excel(source_file, sheet_name=None, header=0)
#creates .csvs from the source file sheets and adds them to the folder
    for sheet_name, data in read_file.items():
        data.to_csv(f"{TODAY}/{sheet_name}.csv", index=False)
    
if __name__ == '__main__':
    main()