#! /usr/bin/env python3
# -*- coding: utf-8 -*-

"""
-------------------     slice_merge_xls_sheet()    ---------------------

This script is used to
   - merge a directory of identically formatted xls/xlsx files
   - extract a window of data from a designated sheet and
   - save the merged data to a csv in the same directory as the program.

A csv file is exported, rather than a spreadsheet, so text tools can be used
to cleanse the file.

The program is run by typing "slice_merge_xls_sheet.py" in the Ancaconda
console. A dialog will appear allowing you to select a directory to import.

Assumptions
1. Only tested in Windows OS
2. All files should contain the same data. Column don't have to be in the
same order but need to have the exact same title. The order of columns will
match the first file imported.

Programmed by Simon Christopher Cropper 28 January 2020

TODO
[ ] document code
[ ] update README

"""

#***********************************************************************
#***********************     GPLv3 License      ************************
#***********************************************************************
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
#***********************************************************************

#-----------------------------------------------------------------------
#--- IMPORTED MODULES
#-----------------------------------------------------------------------

import glob
import sys
import os
import configparser
import pandas as pd

#-----------------------------------------------------------------------
#--- CONFIG
#-----------------------------------------------------------------------

CONFIG = configparser.ConfigParser()
CONFIG.read('slice_merge_xls_sheet.ini')

#-----------------------------------------------------------------------
#--- VARIABLES
#-----------------------------------------------------------------------

DIRNAME = CONFIG['path']['XLS_Dir']
TARGET_SHEET = CONFIG['location']['Sheet']
DROP_COLS = CONFIG['location']['DropCols'].split(",")
DROP_ROWS = CONFIG['location']['DropRows'].split(",")

# Establish where we putting the consolidated
OUTPUT_FILE_NAME = os.getcwd() + '\\output_data\\master'
OUTPUT_FILE_EXT = '.csv'

# Create an output file with sheet suffix
OUTPUT_FILE = OUTPUT_FILE_NAME + "_sheet_" + TARGET_SHEET + OUTPUT_FILE_EXT
TEMP_FILE = OUTPUT_FILE_NAME + "_temp_" + TARGET_SHEET + OUTPUT_FILE_EXT

#-----------------------------------------------------------------------
#--- MAIN PROGRAM
#-----------------------------------------------------------------------


# Continue if dialog returns a directory name
if DIRNAME:
    # Some rudimentary feedback
    print('Selected "{}" directory to import'.format(DIRNAME))

    # Create list of spreadsheets in directory
    FILE_LIST = glob.glob(DIRNAME + "/*.xls?")

    # Establish how many files were collated
    N = len(FILE_LIST)

    # Establish number of sheets in first file
    XLSX_FILE = pd.ExcelFile(FILE_LIST[1])

    # Let the user know what is going on
    print(' ')
    print("There are {} xlsx files in that directory".format(N))

    # Create dataframe store data
    ALL_DATA = pd.DataFrame()

    # Let the user know what is going on
    print(' ')
    print('Extracting {} data from sheets...'.format(TARGET_SHEET))

    # Reiterate through list of files
    for f in FILE_LIST:

        # Create a visual cue to let the user know the program is still importing
        sys.stdout.write('#')
        sys.stdout.flush()

        df = pd.read_excel(f, ignore_index=True, sheet_name=TARGET_SHEET, skiprows=3)
        df.drop(DROP_COLS, axis=1, inplace=True)

        # Transpose Data
        df_transposed = df.transpose()
        df_transposed.to_csv(TEMP_FILE, header=False)
        df_fixed = pd.read_csv(TEMP_FILE, header=0, index_col=0)
        df_fixed.dropna(axis=0, how='all', inplace=True)
        df_fixed['source'] = f

        # Append the single file's data to the consolidated dataframe
        ALL_DATA = ALL_DATA.append(df_fixed, ignore_index=True, sort=False)

    # Importing finished. Save data to CSV
    ALL_DATA.dropna(axis=0, how='all', inplace=True)
    ALL_DATA.drop_duplicates(keep='first', inplace=True)
    ALL_DATA.drop(DROP_ROWS, axis=1, inplace=True)
    ALL_DATA.to_csv(OUTPUT_FILE, index=False)

    # User feedback
    print(" ")
    print('Consolidated {} sheets stored in "{}"'.format(TARGET_SHEET, OUTPUT_FILE))
    os.remove(TEMP_FILE)

# Capture that dialog exited and returns no list
else:

    print("No directory selected. Bye.")
