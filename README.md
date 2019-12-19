# CleanXLSX
VBScript to clean Microsoft Excel workbooks (.xls or .xlsx) that are used with Micro Focus Unified Functional Testing (UFT) tools.

## Purpose

Rarely when using UFT/QTP and Microsoft Excel there could be corruptions with the data. UFT versions prior to 15 do not use Microsoft Excel to interpret, but a third-party tool. This third-party tool does not handle all Excel functionality and can cause issues with data being used in UFT. This tool was developed as an example to clean Excel files before using the files in UFT.

#### Issues addressed
1. Importing data from various sources creates extra blank columns to the right of the data. This can cause an issue where the third-party tool has limit of 256 columns but also takes up memory resources. This VBScript will remove empty columns from the right side until it reaches a column with data.
2. Importing data that has been modified by Microsoft Excel that is not supported by UFT third-party tool. Old versions of UFT(formally QTP) would not work if there was formatting or unsupported functions. This VBScript will strip out the formatting.
3. When importing from HTML pages or other sources, sometimes the page is generated with Non-Breaking spaces. These Non-Breaking spaces look like a space but when doing checkpoints with UFT they are not the same character. This script will look for Non-Breaking spaces (Ascii Decimal 160 Hexadecimal A0) and covert the characters to spaces.
4. When Data Driving UFT will not allow column Headers containing <>~`'?%
5. Check Column headers for duplicate names. If found, it is reported and the duplicate is changed with _DUPE added to end of the column name.

## Overview of Code steps
Cleaning is done by:
1. Saving each sheet as CSV file.
2. Removing data addressed above.
3. Scan headers for characters not allowed and change to underscore
4. Converting all the CSV files to xlsx files.
5. Combining all xlsx files into one new file.
6. Remove temporary files created during the process.
7. Last run information is stored in CleanExcel.log file.

### Usage:

usage: cleanExcel.vbs "<full path to >\source_file" (target) (supress) 

	
 |Attribute | Optional/Required | Details
 | :--------- |:---: |:--------
 |Source file | Required | Path with file name to Excel to be processed enclosed in quotes.
 |Target file | Optional | Name of the new Excel file. (Do not use full path, just file name)
 |Suppress | Optional | If values set to '1' all dialogs will be suppressed.

 #### NOTE
 **Make sure the source file is path is encapsulated with double quotes.**
EXAMPLE: 

 	CleanExcel.vbs "D:\Tom Margrave\GitHub\CleanXLSX\Input.xls"  bob.xls
 
  
 
 ## Acknowledgement
 Thank you to Michael Deveaux for testing and reviewing code.
