#! /usr/bin/env python3
# -*- coding: utf-8 -*-

"""
------------------------     convert_csv_to_xls()    --------------------------

This script is used to merge a directory of csv files to xlsx files of the same
name in the same directory.

The program is run by typing "python convert_csv_to_xls.py" in the Ancaconda
console. A dialog will appear allowing you to select a directory to convert.

Assumptions
- Only tested in Windows OS.
- CSV have headers

Programmed by Simon Christopher Cropper 18 April 2019

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

# Continue if dialog returns a directory name
if ROOT.dirname:

    # Some rudimentary feedbak
    print('Selected "{}" directory to import'.format(ROOT.dirname))

    # Create list of spreadsheets in directory
    FILE_LIST = glob.glob(ROOT.dirname + "/*.csv")

    # Establish how many files were collated
    N = len(FILE_LIST)

    # Let the user know what is going on
    print(' ')
    print("There are {} csv files in that directory".format(N))

    # Reiterate through list of files
    for f in FILE_LIST:

        #-----------------------------------------------------------------------
        #--- CLEANUP NAMES
        #-----------------------------------------------------------------------

        INPUT_PATH, INPUT_FILE_NAME = os.path.split(f)
        INPUT_FILE_NAME, INPUT_FILE_EXTENSION = os.path.splitext(INPUT_FILE_NAME)
        OUTPUT_FILEPATH = INPUT_PATH + '/' + INPUT_FILE_NAME + '.xlsx'

        #-----------------------------------------------------------------------
        #--- FEEDBACK
        #-----------------------------------------------------------------------

        print("Converting {}".format(f))

        #-----------------------------------------------------------------------
        #--- CLEANUP NAMES
        #-----------------------------------------------------------------------

        df = pd.read_csv(f, header=0)

        # Importing finished. Save data to CSV
        df.to_excel(OUTPUT_FILEPATH, index=False)

    # User feedback
    print(" ")
    print('Done')

# Capture that dialog exited and returns no list
else:

    print("No directory selected. Bye.")
