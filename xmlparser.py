# Tecan i-control xml to xlsx parser
# The script takes an xml output from the Tecan i-control software
# and outputs an xlsx file

import os
import re
import sys
import time
import xml.etree.ElementTree as ET
import xlsxwriter

if len(sys.argv) > 1:
    xml = sys.argv[1]
else:
    sys.exit("You need to pass the name of the xml you want to parse, drag and drop is also supported")

filename = os.path.basename(xml)
filename = filename[:-4]
try:
    tree = ET.parse(xml)
except ET.ParseError:
    sys.exit("XML is probably corrupted, please check that there is no tag missing and try again. Exiting.")
except FileNotFoundError:
    sys.exit("File not found. Exiting.")

root = tree.getroot()

workbook = xlsxwriter.Workbook(filename + '.xlsx')
value_format = workbook.add_format({'num_format': '##,####'})
well_format = workbook.add_format({'border': True})
tag_format = workbook.add_format({'bold': True, 'italic': True})
param_format = workbook.add_format({'shrink': True})

duplicate_index = 0 # For when we have name duplicates we need to sort

# As the output is divided in sections, one for each measurement,
# we go from section to section and create a worksheet for each
#
# The XML output already has the Excel positions (EG: "A1")
# it wants the results to be in so we keep track of
# the highest row it writes, so we can safely write after that
for section in root.iter('Section'):
    highestrow = 0

    timestart = section.find('Time_Start').text
    timeend = section.find('Time_End').text
    worksheet_name = section.get('Name')[:30] # Worksheets' names can't be longer than 31
    try:
        worksheet = workbook.add_worksheet(worksheet_name)
    except xlsxwriter.exceptions.DuplicateWorksheetName:
        print("Worksheet '" + worksheet_name + "' already existing, renaming")
        worksheet_name += str(duplicate_index)
        print("Worksheet renamed to " + worksheet_name)
        duplicate_index += 1
        worksheet = workbook.add_worksheet(worksheet_name)
    # Results are divided in "wells"
    for well in section.iter('Well'):
        position = well.get('Pos')
        # We extract the row's number from the Excel position string
        # and convert it to the (0,0) notation
        pos_row = (int(re.search(r'\d+', position).group()) - 1)
        if (highestrow < pos_row): # Keep track of the highest row
            highestrow = pos_row
        # TODO: Change the locale instead of brutally change commas into dots
        value = float((well.find('Single').text).replace(',','.'))
        worksheet.write_number(position, value, well_format)

    # Parameters go into their own worksheet
    worksheet = workbook.add_worksheet("Parameters")
    highestrow = 0
    worksheet.write(highestrow, 0, "Time start:", tag_format)
    worksheet.write(highestrow, 1, timestart)
    highestrow += 1
    worksheet.write(highestrow, 0, "Time end:", tag_format)
    worksheet.write(highestrow, 1, timeend)
    highestrow += 1
    for parameter in section.iter('Parameter'):
        highestrow += 1
        col = 0
        for key in parameter.attrib:
            worksheet.write(highestrow, col, parameter.attrib[key], param_format)
            col += 1

workbook.close()
