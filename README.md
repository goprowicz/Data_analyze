# data-analyze
Data analyzis Python program with batch install file. 
Raw data excel file is turning to processed summary excel file.

Program compares two lines and displays amount of defects for both lines.

Excel files with raw data should be created by template, 
and columns should be arranged in following order:

|detal id|date defect occured|line|defect type|

Avoid blank cells in columns with data, it can cause problems.

Installation:
1. Open install.bat file from install catalogue
2. When Python installator opens, install it.
   IMPORTANT!: "Add python.exe to PATH" box have to be marked.
   Otherwise program may be working incorrectly!

   If you have 32-bit windows download 32-bit windows installer from:
   https://www.python.org/downloads/release/python-3114/ and install it,
   then run install.bat file from this repository, and when installator
   pops out, click cancel. rest of steps will be done automatically
   
   If you have other system than windows download installer for your system,
   and do the same as 32-bit windows case.

Program running:
1. Open run-data_analyze.bat file
2. Window will pop out and you have to browse input data file then,
   click on browse button
3. File explorer will show. Choose excel file you want to analyze.
4. Check if You opened correct file. If so, click on button with filename
5. Another file explorer will open, save your summary file in directory
   you want to keep it, with filename you want.
6. Final excel workbook with processed data will open.
