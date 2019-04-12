# Introduction 

This is a suite of scripts for manipulation of Microsoft Excel files using Python.

The programs are generic so should work in a variety of contexts. 

# Getting Started

These scripts can be run with Python.

In Bupa, this is only possible if you have a Virtual Machine with Anaconda or equalivalent package installed.

# Scripts available

## Merge XLS Files

This script is used to merge a directory of identically formatted xls/xlsx files and save the merged data to a csv in the same directory as the program.

A csv file is exported, rather than a spreadsheet, as some formatting is introduced by the python library if exported to Excel. There is also no opportunity to clean the data before importing into a new master spreadsheet.

The program is run by typing "python Merge_XLS_Files.py" in the Ancaconda console. A dialog will appear allowing you to select a directory to import.

The merged output will be stored in the "output_data" directory for each sheet. This data is best cleansed before continueing.

Assumptions
1.	Only tested in Windows OS
2.	All files should contain the same data. Column don't have to be in the same order but need to have the exact same title. The order of columns will match the first file imported.

# Contribute

If you want to contribute to this list of scripts, clone the VSTS repo and test your updates locally before pushing to master.

Please do not post routines that have hardcoded data references that can't work in most situations.
