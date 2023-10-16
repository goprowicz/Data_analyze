# data-analyze
Python data analyzis program with batch install file. 
Program turns raw excel data file to final look with statistics and pie charts.

Program compares two lines and displays amount of defects for both of them.

Excel files with raw data should be created by template, 
and columns should be arranged in following order:

| detal id | date that defect occured in | line | defect type |

Avoid blank cells in columns with data.

Installation:
1. Open install.bat file from install catalogue
2. When Python installator opens, install it.
   IMPORTANT!: "Add python.exe to PATH" box have to be marked.
   Otherwise program may be working incorrectly!

      If you have 32-bit windows download 32-bit windows installer from:
      https://www.python.org/downloads/release/python-3114/ and install it.
      Then run install.bat file from this repository, and when installator
      pops out, click cancel. rest of steps will be done automatically
      
      If you have other system than windows download installer for your system,
      and do the same as 32-bit windows case.

Program running:
1. Open run-data_analyze.bat file
2. Program will open and you have to browse input data excel file then,
   to do it click on "browse" button.
3. File explorer will open, Choose excel file you want to analyze.
4. Check if You opened correct file. If so, click on button with it's filename
   that will appear after file selection.
6. Another file explorer will open for saving final file.
   Save your summary file in location you want to keep it, with filename you prefer.
8. Final excel workbook with processed data will open.
