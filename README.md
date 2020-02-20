# Tools to manipulate Microsoft Excel files

## Introduction 

This is a suite of scripts for manipulation of Microsoft Excel files using Python.

The programs are generic so should work in a variety of contexts. 

## Getting Started

These scripts can be run with Python 3.

## Scripts available

### Merge XLS Files

This script is used to merge a directory of identically formatted xls/xlsx files and save the merged data to a csv in the same directory as the program.

A csv file is exported, rather than a spreadsheet, as some formatting is introduced by the python library if exported to Excel. There is also no opportunity to clean the data before importing into a new master spreadsheet.

The program is run by typing "python Merge_XLS_Files.py" in the Anaconda console. A dialog will appear allowing you to select a directory to import.

The merged output will be stored in the "output_data" directory for each sheet. This data is best cleansed before continuing.

Assumptions
1.	Only tested in Windows OS
2.	All files should contain the same data. Column don't have to be in the same order but need to have the exact same title. The order of columns will match the first file imported.

### Convert CSV to XLS

This script is used to merge a directory of csv files to xls/xlsx files of the same name in the same directory.

The program is run by typing "python convert_csv_to_xls.py" in the Anaconda console. A dialog will appear allowing you to select a directory to convert.

Assumptions
1. Only tested in Windows OS.
2. CSV have headers

Explanatory Note - The import routine in Excel has limits on the length of fields that can be imported. These limits vary between versions and are significantly smaller than the upper size limit allowed for text cells. These limits do not exist in Python/Pandas.

### Slice then Merge XLS Files

This script is used to
   - merge a directory of identically formatted xls/xlsx files
   - extract a window of data from a designated sheet and
   - save the merged data to a csv in the same directory as the program.

A csv file is exported, rather than a spreadsheet, so text tools can be used to cleanse the file.

The program is run by typing "python slice_merge_xls_sheet.py" in the Anaconda console. A dialog will appear allowing you to select a directory to import. It is essential that a config is placed in the import directory. Sample config files are available in the subdirectory 'sample_ini'. 

The program requires details to be entered in the config file to run. This includes the Sheet to collate,columns to drop and rows to drop.

Assumptions
1. Only tested in Windows OS

## Contribute

If you want to contribute to this list of scripts, clone the VSTS repo and test your updates locally before pushing to master.

Please do not post routines that have hardcoded data references that can't work in most situations.

## Licenses

Software developed under this project is being released under a [GNU General Public License, version 3 (GPL-3.0)][1] license.

Supportive documentation is released under a [Creative Commons Attribution 4.0 International (CC-BY)][2] license.


[1]: https://opensource.org/licenses/GPL-3.0
[2]: https://creativecommons.org/licenses/by/4.0/legalcode

