# OfficeMacros


## Excel Macro: AlignRowsById

Align rows from two spreadsheets with a common ID column.

1. Call the two sheets sheet L and sheet R
1. Sort both spreadsheets by the ID column
1. Insert 2 new columns A and B, which could have headings "L ID" and "R ID"
1. Select all columns from sheet L and insert into sheet R at column A so that the 2 empty columns from the previous step separate the L data from the R data
1. Input a formula into the "L ID" column so that it reflect the ID (or combination key via CONCAT or similar) of the L data
1. Input a formula into the "R ID" column so that it reflect the ID (or combination key via CONCAT or similar) of the R data
1. Run the AlignRowsById macro
1. Within the form, select the "L ID" and "R ID" cells; do NOT select the header (if you added one); do NOT just select the entire column 
1. Click "Realign Rows"
1. It will give you the opportunity to stop every 100 rows for a long spreadsheet


## Getting Started

### Run From Ribbon

1. Download [ExcelMacros](https://raw.githubusercontent.com/brainite/OfficeMacros/master/ExcelMacros.xlsm) and save where it can remain accessible
1. Open ExcelMacros in Excel
1. If you see "SECURITY WARNING: Macros have been disabled.", then click "Enable Content"
1. Keep ExcelMacros open but switch to a new workbook (File > New)
1. In Excel, go to File > Options > Customize Ribbon 
   1. "Choose Commands From:" > "Macros"
   1. Select "ExcelMacros.xslm!AlignRowsById" (adjust for the file name you chose plus the macro you want)
   1. Select the tab and group on the right where you want the option to appear, and click "Add >>"
   1. Click "OK" to close the window
1. In the future, you can use the macros without opening the file first

### Developing OfficeMacros

1. Clone from GitHub
1. Resume the instructions from step 2 above
1. In Excel, go to File > Options > Customize Ribbon 
   1. Enable the "Developer" tab on the right
   1. Click "OK" to close the window
1. Access "Microsoft Visual Basic for Applications" via "Developer" > "View Code"
1. After modifying a macro, right-click on the object in the VBAProject and "Export File..." to prepare it to commit
1. Commit the exported files -- do NOT commit Office document files
1. Issue a pull request
