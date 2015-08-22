# python-excel-stuff

Modules to make working with Excel data easier.


## compare_versions.py

Use: Detecting changes between two versions of an Excel file.<br />

Example: <br />
import compare_versions as cv <br />
cv.compare_excel_files('test_file_1.xlsx', 'test_file_2.xlsx','col1', 'col2') <br />

Output: Excel file with the following worksheets:<br />
        - previous version<br />
        - current version<br />
        - added<br />
        - removed<br />
        - differences (relevant cells highlighted in red)<br />
        
Python version: 2.7<br />

Dependencies: xlrd, xlwt<br />


