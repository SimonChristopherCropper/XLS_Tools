#! /usr/bin/env python3
# -*- coding: utf-8 -*-

"""
-------------------     slice_merge_xls_sheet()    ---------------------

This script is used to
   - merge a directory of identically formatted xls/xlsx files
   - extract a window of data from a designated sheet and
   - save the merged data to a csv in the same directory as the program.

A csv file is exported, rather than a spreadsheet, so text tools can 
be used to cleanse the file.

The program is run by typing "python slice_merge_xls_sheet.py" in 
the Ancaconda console. 

The program requires details to be entered in the config file to run.
This includes the Sheet to collate,columns to drop and rows to drop.

Assumptions
1. Only tested in Windows OS

Programmed by Simon Christopher Cropper 28 January 2020

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
import tkinter as tk
from tkinter import filedialog

#-----------------------------------------------------------------------
#--- DIRECTORY DIALOG
#-----------------------------------------------------------------------

ROOT = tk.Tk()
ROOT.withdraw()
ROOT.dirname = filedialog.askdirectory()

#-----------------------------------------------------------------------
#--- BASIC CHOICES
#-----------------------------------------------------------------------

if ROOT.dirname:
    ini_file = ROOT.dirname + os.sep + 'slice_merge_xls_sheet.ini'
    if os.path.exists(ini_file):
        Proceed = True
    else:
        print("No ini file in directory. Bye.")
        Proceed = False
else:
    print("No directory selected. Bye.")
    Proceed = False

#-----------------------------------------------------------------------
#--- MAIN PROGRAM
#-----------------------------------------------------------------------

# Continue if dialog returns a directory name
if Proceed:

    #-----------------------------------------------------------------------
    #--- CONFIG
    #-----------------------------------------------------------------------

    CONFIG = configparser.ConfigParser()
    CONFIG.read(ini_file)

    #-----------------------------------------------------------------------
    #--- VARIABLES
    #-----------------------------------------------------------------------

    DIRNAME = ROOT.dirname
    TARGET_SHEET = CONFIG['location']['Sheet']
    LINES_HEADER = int(CONFIG['location']['Lines2Header'])
    DROP_COLS = CONFIG['location']['DropCols'].split(",")
    DROP_ROWS = CONFIG['location']['DropRows'].split(",")

    # Establish where we putting the consolidated
    OUTPUT_FILE_NAME = os.getcwd() + '\\output_data\\master'
    OUTPUT_FILE_EXT = '.csv'

    # Create an output file with sheet suffix
    OUTPUT_FILE = OUTPUT_FILE_NAME + "_sheet_" + TARGET_SHEET + OUTPUT_FILE_EXT
    TEMP_FILE = OUTPUT_FILE_NAME + "_temp_" + TARGET_SHEET + OUTPUT_FILE_EXT

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

        # Reads in sheet from workbook, starting from the line after Lines2Header
        df = pd.read_excel(f, ignore_index=True, sheet_name=TARGET_SHEET, skiprows=LINES_HEADER)
        df.drop(DROP_COLS, axis=1, inplace=True)

        # Transpose Data, pushed through csv to force column and row labels to stick
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

