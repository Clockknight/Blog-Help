import os
import sys
import openpyxl
from openpyxl import workbook

content = ''

#.xlsx argv check
if sys.argv[1][-5:] != '.xlsx' or len(sys.argv) != 2:
    print("Apologies. The argument provided did not lead to a valid .xlsx file. Please try again.")
    sys.exit()
#os.exists check

#Create variables based on spreadsheet

#Store data on content variable

#Save
