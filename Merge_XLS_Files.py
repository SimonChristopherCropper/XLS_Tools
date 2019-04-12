#! /usr/bin/env python3
# -*- coding: utf-8 -*-

"""
------------------------     Merge_XLS_Files()    --------------------------

This script is used to merge a directory of identically formatted xls/xlsx
files and save the merged data to a csv in the same directory as the program.

A csv file is exported, rather than a spreadsheet, as some formatting
is introduced by the python library.

The program is run by typing "python Merge_XLS_Files.py" in the Ancaconda
console. A dialog will appear allowing you to select a directory to import.

Assumptions
1. Only tested in Windows OS
2. All files should contain the same data. Column don't have to be in the 
same order but need to have the exact same title. The order of columns will 
match the first file imported.


Programmed by Simon Christopher Cropper 10 April 2018

"""

#-----------------------------------------------------------------------
#--- IMPORTED MODULES
#-----------------------------------------------------------------------

import pandas as pd
import numpy as np
import glob
import sys
import os
from tkinter import filedialog
from tkinter import *

#-----------------------------------------------------------------------
#--- DIRECTORY DIALOG
#-----------------------------------------------------------------------

root = Tk()
root.withdraw()
root.dirname =  filedialog.askdirectory()

#-----------------------------------------------------------------------
#--- MAIN PROGRAM
#-----------------------------------------------------------------------

# Establish where we putting the consolidated
output_file_name = os.getcwd() + '\output_data\master'
output_file_ext = '.csv'

# Continue if dialog returns a list of names
if root.dirname:

	# Some rudimentary feedbak
	print('Selected "{}" directory to import'.format(root.dirname)) 
	
	# Create dataframe store data
	all_data = pd.DataFrame()
	
	# Rudimentary trick to ensure headerf is only collected once
	one_time = True
	
	# Create list of spreadsheets in directory
	filelist = glob.glob(root.dirname + "/*.xls?")
	
	# Establish how many files were collated
	n = len(filelist)
	
	# Establish number of sheets in first file
	xlsx_file = pd.ExcelFile(filelist[1])
	NumSheets = len(xlsx_file.sheet_names)
	
	# Let the user know what is going on
	print(' ')
	print("There are {} xlsx files in that directory".format(n))
	
	for SheetIndex in range(0, NumSheets):
		
		# Create an output file with sheet suffix
		SheetLabel = str(SheetIndex + 1)
		output_file = output_file_name + "_sheet_" + SheetLabel + output_file_ext
		
		# Let the user know what is going on
		print(' ')
		print('Importing data from Sheet {}...'.format(SheetLabel))
		
		# Reiterate through list of files
		for f in filelist:
			
			# Create a visual cue to let the user know the program is still importing
			sys.stdout.write('#')
			sys.stdout.flush()
			
			# Record the header of the first file imported
			if one_time :
				df = pd.read_excel(f,ignore_index=True, sheet_name=SheetIndex) 
				one_time = False
			# Ignore the header on the remaining files
			else :
				df = pd.read_excel(f,ignore_index=True,sheet_name=SheetIndex,skip_row=0) 
			
			# Append the single file's data to the consolidated dataframe
			all_data = all_data.append(df,ignore_index=True, sort=False) 

		# Importing finished. Save data to CSV
		all_data.to_csv(output_file, index=False)
		
		# User feedback
		print(" ")
		print('Data in sheet {} stored in "{}"'.format(SheetLabel, output_file))

# Capture that dialog exited and returns no list
else:

	print("No directory selected. Bye.")