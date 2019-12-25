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
book = openpyxl.load_workbook(str(sys.argv[1]))
sheet = book.active
    #Height variable
height = sheet.max_row
print(str(height - 1) + ' items found. Processing...')


#Store data on content variable
content += '''[su_table class="custom-su-table" responsive="yes"]
<table>
<tr class="header">
<td>Box Art</td>
<td>Game</td>
<td>Player(s)</td>
<td>Playtime</td>
<td>Rating</td>
</tr>'''

for i in range(2, height+1):
    #Take href value from column 1, add it to an a tag with column 2's info
    nameString = '^^^'

    #Write values from each column
    item = '\n\n<tr>'
    for j in range(1, 5):
        body = '\n<td>' + str(sheet.cell(row=i, column=j).value) + '</td>'
        item += body + '\n</tr>'

    #item += sheet.cell(row=i, column=1).value + '</td>\n<td>' + nameString + '</td>\n<td>' + sheet.cell(row=i, column=3).value + '</td>\n<td>' + sheet.cell(row=i, column=4).value + '</td>\n<td>' + sheet.cell(row=i, column=5).value + '</td>\n</tr>'
    content += item
#End the su_table tag
content += '\n[/su_table]'

#Write out value
file = open('blogHelp.txt', 'w+')
file.write(content)
file.close()

#Save
