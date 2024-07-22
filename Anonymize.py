# -*- coding: utf-8 -*-
"""
Created on Mon Jul 22 11:17:33 2024

@author: Hadrien

The goal for this script is to scramble and anonymize data by replacing names by
randomly generated strings.
"""

import random
import string
import pandas as pd

#input name of default source file here
DEFAULT_SOURCE = "ExcelSource.xlsx"

def main():
    read()
    anonymize()

def read():
    sourcefile = DEFAULT_SOURCE
    target = pd.read_excel(sourcefile, sheet_name=0, header=0, usecols="A")
    return target
    
def newname():
    letters = string.ascii_letters
    newname = ( ''.join(random.choice(letters) for i in range(10)) )
    return newname

def anonymize():
    for row in target


if __name__ == '__main__':
    main()