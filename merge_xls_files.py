#! /usr/bin/env python3
# -*- coding: utf-8 -*-

"""
------------------------     merge_xls_files()    --------------------------

This script is used to merge a directory of identically formatted xls/xlsx
files and save the merged data to a csv in the same directory as the program.

A csv file is exported, rather than a spreadsheet, as some formatting
is introduced by the python library.

The program is run by typing "python merge_xls_files.py" in the Ancaconda
console. A dialog will appear allowing you to select a directory to import.

Assumptions
1. Only tested in Windows OS
2. All files should contain the same data. Column don't have to be in the
same order but need to have the exact same title. The order of columns will
match the first file imported.

Programmed by Simon Christopher Cropper 10 April 2019

This is a new line

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
import tkinter as tk
from tkinter import filedialog
import pandas as pd

#-----------------------------------------------------------------------
#--- DIRECTORY DIALOG
#-----------------------------------------------------------------------

ROOT = tk.Tk()
ROOT.withdraw()
ROOT.dirname = filedialog.askdirectory()

#-----------------------------------------------------------------------
#--- MAIN PROGRAM
#-----------------------------------------------------------------------

# Establish where we putting the consolidated
OUTPUT_FILE_NAME = os.getcwd() + '\\output_data\\master'
OUTPUT_FILE_EXT = '.csv'

# Continue if dialog returns a directory name
if ROOT.dirname:

    # Some rudimentary feedbak
    print('Selected "{}" directory to import'.format(ROOT.dirname))

    # Create list of spreadsheets in directory
    FILE_LIST = glob.glob(ROOT.dirname + "/*.xls?")

    # Establish how many files were collated
    N = len(FILE_LIST)

    # Establish number of sheets in first file
    XLSX_FILE = pd.ExcelFile(FILE_LIST[1])
    NUM_SHEETS = len(XLSX_FILE.sheet_names)

    # Let the user know what is going on
    print(' ')
    print("There are {} xlsx files in that directory".format(N))

    for SheetIndex in range(0, NUM_SHEETS):

        # Create dataframe store data
        ALL_DATA = pd.DataFrame()

        # Create an output file with sheet suffix
        SheetLabel = str(SheetIndex + 1)
        output_file = OUTPUT_FILE_NAME + "_sheet_" + SheetLabel + OUTPUT_FILE_EXT

        # Let the user know what is going on
        print(' ')
        print('Importing data from Sheet {}...'.format(SheetLabel))

        # Reiterate through list of files
        for f in FILE_LIST:

            # Create a visual cue to let the user know the program is still importing
            sys.stdout.write('#')
            sys.stdout.flush()

            # Read data into dataframe
            df = pd.read_excel(f, sheet_name=SheetIndex)

            # Append the single file's data to the consolidated dataframe
            ALL_DATA = ALL_DATA.append(df, sort=False)

        # Importing finished. Save data to CSV
        ALL_DATA.to_csv(output_file, index=False)
 
        # User feedback
        print(" ")
        print('Data in sheet {} stored in "{}"'.format(SheetLabel, output_file))

# Capture that dialog exited and returns no list
else:

    print("No directory selected. Bye.")
