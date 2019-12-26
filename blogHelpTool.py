import os
import sys
import bs4
from bs4 import BeautifulSoup
import openpyxl
from openpyxl import workbook

#.xlsx argv check
if sys.argv[1][-5:] != '.xlsx' or len(sys.argv) != 2:
    print("Apologies. The argument provided did not lead to a valid .xlsx file. Please try again.")
    sys.exit()
#os.exists check


#Store data on spreadsheetText variable
spreadsheetText = '''[su_table class="custom-su-table" responsive="yes"]
<table>
<tr class="header">
<td>Box Art</td>
<td>Game</td>
<td>Player(s)</td>
<td>Playtime</td>
<td>Rating</td>
</tr>'''
linkHeaders = ''
#Create variables based on spreadsheet
book = openpyxl.load_workbook(str(sys.argv[1]))
sheet = book.active
    #Height variable
height = sheet.max_row
print(str(height - 1) + ' items found. Processing...')

for i in range(2, height+1):
    #Take a tag from column 1's data, add it to an a tag with column 2's info
    itemHtml = str(sheet.cell(row=i, column=1).value)
    soup = BeautifulSoup(itemHtml, 'html.parser')
    gameaTag = str(soup.a)

    #Write values from each column
    itemDetails = '\n\n<tr>'
    for j in range(1, 6):
        cellInfo = '\n<td>' + str(sheet.cell(row=i, column=j).value) + '</td>'

        if j == 2:  #Special clause for 2nd column's data
            cellInfo = '\n<td>' + gameaTag + str(sheet.cell(row=i, column=2).value) + '</a></td>'
            linkHeaders = linkHeaders+ '<h3>' + gameaTag  + '</h3>\n'

        #Write completed cellInfo to itemDetails
        itemDetails += cellInfo

    spreadsheetText += itemDetails + '\n</tr>'

spreadsheetText += '\n[/su_table]\n\n'
spreadsheetText += linkHeaders

#Write out value
file = open('blogHelp.txt', 'w+')
file.write(spreadsheetText)
file.close()
