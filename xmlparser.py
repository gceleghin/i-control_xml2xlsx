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

def draw_plate_fluorescence(worksheet, cycle, rows_number = 8, columns_number = 12,):
    row = (11 * ((cycle) - 1)) + 1

    for col_num in range(columns_number):
        worksheet.write(row, col_num + 1, col_num + 1, plate_format)

    for row_num in range(rows_number):
        worksheet.write(row + 1, 0, string.ascii_uppercase[row_num], plate_format)
        row += 1

def draw_plate_scan(worksheet, dataset, cycle, wavelength_start, wavelength_end, wavelength_step):
    row = (11 * ((cycle) - 1)) + 1
    columns_number = int(((wavelength_end - wavelength_start) / wavelength_step) + 1)

    for col_num in range(columns_number):
        wavelength = wavelength_start + (wavelength_step * (col_num))
        worksheet.write(row, col_num + 1, wavelength, plate_format)

    for well in dataset.iter('Well'):
        position = well.get('Pos')
        pos_col = 0
        row += 1
        worksheet.write(row, pos_col, position, plate_format)
        for scan in well.iter('Scan'):
            wave = float(scan.get('WL'))
            pos_col = ((wave - wavelength_start) / wavelength_step) + 1
            value = float(scan.text)
            worksheet.write_number(row, pos_col, value, measured_well_format)

    write_parameters(worksheet, columns_number + 2)

def write_fluorescence_data(worksheet_to_draw, dataset, cycle):
    # Inside each cycle, each measurement is in a <Well> tag
    for well in dataset.iter('Well'):
        position = well.get('Pos')
        # Extracts the numbers in the position
        pos_column = int(re.search(r'\d+', position).group())
        # Gets the letter in the position
        pos_row = string.ascii_uppercase.index(position[0])
        # We multiply so subsequent cycles don't overwrite each other
        pos_row = (11 * (cycle - 1)) + pos_row + 2
        # TODO: Change the locale instead of brutally change commas into dots
        value = float((well.find('Single').text).replace(',','.'))
        status = well.find('Single').get('Status')
        if status == "Invalid":
            worksheet.write_number(pos_row, pos_column, value, invalid_well_format)
        else:
            worksheet.write_number(pos_row, pos_column, value, measured_well_format)

    write_parameters(worksheet, 15)

def write_parameters(worksheet, start_column, current_row = 0):
    # Section parameters go on the right
    worksheet.set_column(start_column, start_column + 1, 25)
    time_start = section.find('Time_Start').text
    time_end = section.find('Time_End').text
    worksheet.write(current_row, start_column, "Time start:", tag_format)
    worksheet.write(current_row, start_column + 1, time_start)
    current_row += 1
    worksheet.write(current_row, start_column, "Time end:", tag_format)
    worksheet.write(current_row, start_column + 1, time_end)
    current_row += 1
    for parameter in section.iter('Parameter'):
        current_row += 1
        column = start_column
        for key in parameter.attrib:
            worksheet.write(current_row, column, parameter.attrib[key], param_format)
            column += 1

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
measured_well_format = workbook.add_format({'border': True})
invalid_well_format = workbook.add_format({'border': True, 'bg_color': 'yellow'})
tag_format = workbook.add_format({'bold': True, 'italic': True})
param_format = workbook.add_format({'shrink': True})
plate_format = workbook.add_format({
    'bold': True, 'font_color': 'white', 'bg_color': '008080', 'align': 'center'})
cycle_info_format = workbook.add_format({'font_color': 'white', 'bg_color': '008080', 'align': 'left'})

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

    worksheet.set_column(0, 0, 3)

    # Each section is then divided in cycles
    for dataset in section.iter('Data'):
        cycle = int(dataset.attrib["Cycle"])

        cycle_first_row = (11 * (cycle - 1))
        first_column = 1
        worksheet.merge_range(cycle_first_row, first_column, cycle_first_row, first_column + 1, "Cycle:", cycle_info_format)
        first_column += 2
        worksheet.merge_range(cycle_first_row, first_column, cycle_first_row, first_column + 1, cycle, cycle_info_format)

        try:
            cycle_start = dataset.attrib["Time_Start"]
            first_column += 2
            worksheet.merge_range(cycle_first_row, first_column, cycle_first_row, first_column + 1, "Time start:", cycle_info_format)
            first_column += 2
            worksheet.merge_range(cycle_first_row, first_column, cycle_first_row, first_column + 1, cycle_start, cycle_info_format)
        except KeyError as e:
            pass

        try:
            cycle_temperature = dataset.attrib['Temperature']
            first_column += 2
            worksheet.merge_range(cycle_first_row, first_column, cycle_first_row, first_column + 1, "Temperature:", cycle_info_format)
            first_column += 2
            worksheet.merge_range(cycle_first_row, first_column, cycle_first_row, first_column + 1, cycle_temperature, cycle_info_format)
        except KeyError as e:
            pass

        if first_column < 11:
            first_column += 2
            worksheet.merge_range(cycle_first_row, first_column, cycle_first_row, 12, "", cycle_info_format)

        try:
            wavelength_start = int(section.find("./Parameters/Parameter[@Name='Emission Wavelength Start']").attrib['Value'])
            wavelength_end = int(section.find("./Parameters/Parameter[@Name='Emission Wavelength End']").attrib['Value'])
            wavelength_step = int(section.find("./Parameters/Parameter[@Name='Emission Wavelength Step Size']").attrib['Value'])
            draw_plate_scan(worksheet, dataset, cycle, wavelength_start, wavelength_end, wavelength_step)
        except AttributeError as e:
            draw_plate_fluorescence(worksheet, cycle)
            write_fluorescence_data(worksheet, dataset, cycle)

try:
    workbook.close()
except xlsxwriter.exceptions.FileCreateError as e:
    print("Can't create file. It may happen if you have the file already open or",
        "if you don't have write permission in the folder you are running the script in.")
    print(e)
