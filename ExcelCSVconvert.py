# -*- coding: utf-8 -*-
"""
Created on Fri Jun  7 12:06:56 2024

@author: Hadrien

This script serves to extract each sheet of a source Excel file to CSVs.
The goal is to use this as one step in a workflow where data is prepared
in Excel then exported into a database via CSVs. 
"""
import pandas as pd
from datetime import date
import os

#input name of default source file here
DEFAULT_SOURCE = "ExcelSource.xlsx"

#set date to track when files were made
DEFAULT_FOLDER = str(date.today())

def main():
    intro()
    while True:
        write(extract(), folder())
        if repeat_check() == False:
            break
    outro()

#verify the source then extract it to the folder
def extract():
    sourcefile = DEFAULT_SOURCE
    print("The file to be processed is set by default to: " + sourcefile + "\n")
    print("Is this the file you want to process? (Y/N) ")
    #strip and upper the input to minimize error
    answer = str(input()).strip().upper()
    while (answer != "N") and (answer != "Y"):
        answer = str(input("Please confirm 'Y' or 'N'. ")).strip().upper()
    if answer == "N":
        print("\nEnter the name of the source file including the .extension: ")   
        sourcefile = input()
        #check whether the source can be found to prevent FileNotFoundErrors
        while not os.path.exists(sourcefile):
            print("\nThat file cannot be found in the folder. Please enter the name again: ")
            sourcefile = str(input()).strip()
    elif answer == "Y":
            while not os.path.exists(sourcefile):
                print("\nThat file cannot be found in the folder. Please add it and run the script again or enter a new file source: ")
                sourcefile = str(input()).strip()
    print("\nThe script will now extract the sheets from " + sourcefile)
    return sourcefile

#create a new folder with the current date if it doesnt exist yet
def folder():
    folderpath = DEFAULT_FOLDER
    print("\nBy default, the script will extract files to a folder named: " + folderpath + ".\nIs this OK? (Y/N)")
    answer = str(input()).strip().upper()
    while (answer != "N") and (answer != "Y"):
        answer = str(input("\nPlease confirm 'Y' or 'N'. "))
    if answer == "N":
        print("\nEnter the name of the folder: ")   
        folderpath = input()
    if not os.path.exists(folderpath):
        os.makedirs(folderpath)
    return folderpath

#combine the extract and folder setting to create the csvs
def write(sourcefile, folderpath):    
    #reads for the source Excel
    read_file = pd.read_excel(sourcefile, sheet_name=None, header=0)
    #creates .csvs from the source file sheets and adds them to the folder
    for sheet_name, data in read_file.items():
        data.to_csv(f"{folderpath}/{sheet_name}.csv", index=False)
        
#sets conditions for the script to loop or end
def repeat_check():
    print("Do you need to run the script again? (Y/N)")
    repeatcheck = str(input()).strip().upper()  
    if repeatcheck == "N":
        print("Understood. The script will close.")
        return False
    else:
        print("Understood. The script will now repeat.")
    print()

#defining the text appearing at the start and end of the script to introduce/confirm
def intro():
    print("Hello friend user!")
    print("This script will export individual sheets from the source Excel to their own .csv files.")
    print()

def outro():
    print("Extract complete. Have a nice day!")

if __name__ == '__main__':
    main()
