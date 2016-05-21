## Websense Policy tool

The point of this tool is to make CSV files from a websense excel policy file.
The data in the excel policy spreadsheet doesn't make sense. If we output it in different CSV files that make sense to read then the engineers can make use of it.

The xls files from websense can't be parsed by XLRD since they are html
table files which are openable in Excel. The workaround is to open the xls files in Open Office or MS Office and save them as xlsx files. XLRD can parse and use xlsx files with no problem.



### Dependencies:

$ sudo pip install xlrd


## Usage:

Use the flags:
-f : Input File
-s : The category name
-o : Output File name

eg: python wsParser.py -f testdata/policies.xlsx -s "Category Filter" -o testdata/category_out.csv
