This repository is where I put scripts inspired by my struggles manually wrangling data. 
The included scripts are as follows:

- [Anonymize](https://github.com/paxhadriana/Data-Tools#anonymize): Preparation stage
- [Excel-CSV convert](https://github.com/paxhadriana/Data-Tools#excel-csv-convert): Basic functionality implemented

# Anonymize
## Objective
This script aims to scramble and anonymize data by replacing names in an Excel sheet with randomly generated strings.

A possible use case is when internal documents are shared with external parties as part of information gathering, scope definition or requirements collection.

## To-Do
- [ ] Basic core functionality: replace all name strings in a column with randomly generated strings
- [ ] Allow user to specify which column(s) need to be anonymized
- [ ] Maintain groups by randomizing a given name to the same result if it is found again to maintain aggregation possibilities

# Excel-CSV convert
## Objective
This script is used to extract .csv files from an Excel document's sheets. 

The intended use case is as part of processes which move data from Excel to relational databases which ingest .csvs better.

## To-Do
- [x] Allow user input to change default/hardcoded source file and destination folder
- [x] Allow user to extract from several sources (**Partial**: script can now loop and run through sources in succession without ending unless confirmed)
- [ ] Add some automated error catching (eg. if inputted source does not contain file extension, try accepted ones before invalidating)
- [ ] Allow user to select specific sheets to extract
- [ ] Allow user to target specific columns within a sheet to extract
