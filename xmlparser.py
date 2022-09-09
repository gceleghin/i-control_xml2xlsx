#!/usr/bin/env python3

# Tecan i-control xml to xlsx parser
# The script takes an xml output from the Tecan i-control software
# and outputs an xlsx file

import os
import re
import sys
import time
import string
import xml.etree.ElementTree as ET
import xlsxwriter

def draw_plate(worksheetToDraw, cycleToDraw, rowsNumber = 8, columnsNumber = 12,):
    row = 11 * ((cycleToDraw) - 1)

    for colNum in range(columnsNumber):
        worksheetToDraw.write(row, colNum + 1, colNum + 1, plateFormat)

    for rowNum in range(rowsNumber):
        worksheetToDraw.write(row + 1, 0, string.ascii_uppercase[rowNum], plateFormat)
        row += 1

if len(sys.argv) > 1:
    xml = sys.argv[1]
else:
    sys.exit(
        "You need to pass the name of the xml you want to parse, drag and drop is also supported"
    )

filename = os.path.basename(xml)
filename = filename[:-4]
try:
    tree = ET.parse(xml)
except ET.ParseError:
    sys.exit(
        "XML is probably corrupted, please check that there is no tag missing and try again. Exiting."
    )
except FileNotFoundError:
    sys.exit("File not found. Exiting.")

root = tree.getroot()

workbook = xlsxwriter.Workbook(filename + '.xlsx')

value_format = workbook.add_format({'num_format': '##,####'})
well_format = workbook.add_format({'border': True})
tag_format = workbook.add_format({'bold': True, 'italic': True})
param_format = workbook.add_format({'shrink': True})
plateFormat = workbook.add_format({
    'bold': True, 'font_color': 'white', 'bg_color': 'black', 'align': 'center'
})

duplicate_index = 0 # For when we have name duplicates we need to sort

# As the output is divided in sections, one for each measurement,
# we go from section to section and create a worksheet for each

for section in root.iter('Section'):

    # Worksheets' names can't be longer than 31, leaving some space for additions
    worksheet_name = section.get('Name')[:24]

    try:
        worksheet = workbook.add_worksheet(worksheet_name)
    except xlsxwriter.exceptions.DuplicateWorksheetName:
        print("Worksheet '" + worksheet_name + "' already existing, renaming")
        worksheet_name += str(duplicate_index)
        print("Worksheet renamed to " + worksheet_name)
        duplicate_index += 1
        worksheet = workbook.add_worksheet(worksheet_name)

    # Each section is then divided in cycles
    for dataset in section.iter('Data'):
        cycle = int(dataset.attrib["Cycle"])

        draw_plate(worksheet, cycle)

        # Inside each cycle, each measurement is in a <Well> tag
        for well in dataset.iter('Well'):
            position = well.get('Pos')
            # Extracts the numbers in the position
            posColumn = int(re.search(r'\d+', position).group())
            # Gets the letter in the position
            posRow = string.ascii_uppercase.index(position[0])
            # We multiply so subsequent cycles don't overwrite each other
            posRow = (11 * (cycle - 1)) + posRow + 1
            # TODO: Change the locale instead of brutally change commas into dots
            value = float((well.find('Single').text).replace(',','.'))
            worksheet.write_number(posRow, posColumn, value, well_format)

    # Section parameters go into their own worksheet after the cycles
    worksheet = workbook.add_worksheet(worksheet_name + "_param")
    worksheet.set_column(0, 1, 25)
    highestRow = 0
    timestart = section.find('Time_Start').text
    timeend = section.find('Time_End').text
    worksheet.write(highestRow, 0, "Time start:", tag_format)
    worksheet.write(highestRow, 1, timestart)
    highestRow += 1
    worksheet.write(highestRow, 0, "Time end:", tag_format)
    worksheet.write(highestRow, 1, timeend)
    highestRow += 1
    for parameter in section.iter('Parameter'):
        highestRow += 1
        col = 0
        for key in parameter.attrib:
            worksheet.write(highestRow, col, parameter.attrib[key], param_format)
            col += 1

try:
    workbook.close()
except xlsxwriter.exceptions.FileCreateError as e:
    print("Can't create file. It may happen if you have the file already open or",
        "if you don't have write permission in the folder you are running the script in.")
    print(e)
