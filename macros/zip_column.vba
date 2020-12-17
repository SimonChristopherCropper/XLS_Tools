Sub Zip_Column()
'
'-------------------------- Zip_Column() ------------------------------ 
'
' This macro will create a comma separated list of all the values from 
' the column below the cell selected. Blanks will be ignored. The 
' concatenated cleaned list of values will be inserted in the cell 
' two cell below the last value in a column.
'
' None destructive collapsing of a column of strings or comma separated 
' strings into a single field. Duplicates are removed and the 
' data is sorted.
'
' Assumptions: 
' (1) There will not be more than 10K cells being condensed; 
' (2) there will be less than 1000 unique values
' (3) the standard delimiter is a comma and the output delimiter a 
'     semi-colon; 
' (4) routine is case insensitive - so "aaa" is the same 
'     as "AAA"
'
' Programmed by Simon Christopher Cropper 17 December 2020
'
'***********************************************************************
'***********************     GPLv3 License      ************************
'***********************************************************************
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'***********************************************************************

    ' Create some basic variables
    Dim myCells As Range
    Dim firstCell As Range
    Dim lastCell As Range
    Comma0 = ""","""
    Comma1 = """, """
    Comma2 = """ ,"""
    FinalDelimiter = """; """
  
    ' Record where we are
    Set firstCell = Selection
    
    ' Record last cell in column with a value. This techniques allows 
    'blanks to be skipped
    Set lastCell = Cells(Rows.Count, firstCell.Column).End(xlUp)
    
    ' Select range and copy
    Range(firstCell.Address, lastCell.Address).Select
    Selection.Copy
    
    ' Figure out where we want the final list
    Range(lastCell.Address).Select
    ActiveCell.Offset(2, 0).Select
    Set ShortList = Selection
   
    ' Go to a working area within the same sheet
    Range("A10000").Select
    
    ' Put working copy of data here
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' remove duplicates
    Selection.RemoveDuplicates Columns:=1, Header:=xlNo
    
    ' Flip the data so blanks are at bottom so we can identify only 
    ' those cells with values
    Set myCells = Selection
    Application.Range(myCells.Address).SortSpecial Order1:=xlDescending

    ' Select values for list (currently reversed, i.e. z-a order)
    Selection(1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Set myCells = Selection
    
    ' Clean up the list. Assumes cells may actually have a comma separated
    ' list of values.    
    Selection.End(xlDown).Select
    ActiveCell.Offset(2, 0).Select
    ActiveCell.Formula = "=SUBSTITUTE(SUBSTITUTE(TEXTJOIN(" & Comma1 & ",TRUE," & myCells.Address & ")," & Comma1 & "," & Comma0 & ")," & Comma2 & "," & Comma0 & ")"
    Selection.Copy
    
    ' Move focus and paste values
    ActiveCell.Offset(2, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    ' Split data horizontally
    Selection.TextToColumns Destination:=ActiveCell, DataType:=xlDelimited, ConsecutiveDelimiter:=False, Comma:=True, Space:=False
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
    
    ' dump cell range into a column
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True

    ' Remove duplicates resort list in a-z format
    Selection.RemoveDuplicates Columns:=1, Header:=xlNo
    Selection(1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Set myCells = Selection
    Application.Range(myCells.Address).SortSpecial Order1:=xlAscending

    ' Work out the place to put formula and insert
    Selection.End(xlDown).Select
    ActiveCell.Offset(2, 0).Select
    ActiveCell.Formula = "=TEXTJOIN(" & FinalDelimiter & ",TRUE," & myCells.Address & ")"

    ' Copy cell with list (currently a formula)
    Selection.Copy

    ' Return to cell just below original list, past values
    ShortList.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' Clean away the working values
    Rows("10000:11000").Select
    Selection.Delete Shift:=xlUp
    ShortList.Select

End Sub


