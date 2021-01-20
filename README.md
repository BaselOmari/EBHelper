# Medical-Research-Source-Organizer

Welcome to the Medical Research Source Organzer

Project retrieves relevant scholar articles and publications on the EBSCOhost database and records citation details to help simplify the referencing process. Program was used by University of Nicosia medical students for research on ‘the use of EEG devices to aid anesthesiologists during operations.’


## Project Structure
- [EBSCO Data Collection]("https://github.com/BaselOmari/EBHelper/blob/main/EBSCO%20Data%20Collection.py"): Data from the EBSCOhost database are collected for a specific search are collected and stored in CSV files
- [Excel Sheet Organization]("https://github.com/BaselOmari/EBHelper/blob/main/Excel%20Sheet%20Organizer.py"): Using the Panda Library, the stored CSV files are read and data (Title, Author, Abstract, Publication Date) is neatly organized in an excel spreadsheet.

## Follow Up Work
Students required help with organizing their excel sheets. Files included in the [Follow Up Directory]("https://github.com/BaselOmari/EBHelper/tree/main/Follow%20Up%20Work") are those used to assist the students:
- [Duplicate Removing]("https://github.com/BaselOmari/EBHelper/blob/main/Follow%20Up%20Work/DuplicateRemoving.py"): Since data was extracted from both the EBSCO and Pubmed Database, some duplicates existed between both sources. Script removes duplicates based on which one has more data, and reorganizes the spreadsheet accordingly.
- [Arrangement Fix]("https://github.com/BaselOmari/EBHelper/blob/main/Follow%20Up%20Work/ArrangementFix.py"): Students attempted to arrange cells in Alphabetical Order which disorganized the assisting citation data. Script collects data for each title and reorganizes the spreadsheet accordingly.
