#! python
# this script can invert cells from all worksheets of current workbook
# Instructions:
# type: python cellInverter.py [filename]
# script will create a new workbook with modified name of original file
# X 2020 Arnold Cytrowski

import os, sys, openpyxl

if len(sys.argv) > 2:
    print('try again: [python cellInverter.py [properfilename]]')
    exit()

wb = openpyxl.open(sys.argv[1])
oldsheet = wb.active

output = openpyxl.Workbook()
newsheet = output.active

('Inverting the cells...')

for x in range(1, oldsheet.max_row + 1):
    for y in range(1, oldsheet.max_column + 1):
        newsheet.cell(row = y, column = x).value = oldsheet.cell(row = x, column = y).value

output.save(f'Inverted{os.path.basename(sys.argv[1])}')

print('aaand it\'s done')




