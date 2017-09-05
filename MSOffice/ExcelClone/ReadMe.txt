NOTE: THIS VERSION HAS BEEN DEPRECATED AND REPLACED WITH THE 'ExcelClone_ColumnPick'. PLEASE USE LINK BLOW FOR REFERENCE. THIS NEW CODE CREATES A MORE STABLE FULL CLONE AS WELL AS ALLOWS COLUMN PICKING. 
https://github.com/sayuga/VBA-Tools/tree/master/MSOffice/ExcelClone_ColumnPick

***********************************************************************************************************************************
===Excel Cloner VBA===

Developer: Jonathan Vargas 
Version: 1.0.0
Revision Date: 12-14-2016
Requires: Macro-Enabled Excel File
Source Range Size Limit: A1:BZ5000
Copyright: MIT/X11

Create douplicate excel report with selected columns from source file using a user interface.  

===Description===

This macro brings up a User Interface (UI) which allows you to douplicate selective columns from the source file. 

Even if the source data is fed as a protected sheet, this macro will copy all entries within the A1:BZ5000 cell range and output selected columns in a secondary Excel sheet. 

== Changelog ==

= 0.9.1 =
* Output full copy on secondary excel

= 0.9.2 =
* Added User Interface (UI)
* Added Individual Culumns pickers
* Added 'Select All' picker
* Added Report Name Designation
* Added Sheet Source Name Designation
* Added protected Watermark to first cell designating 'COPY_ONLY_' status

= 0.9.3 =
* Updated UI to close on Submit.
* Created README.txt file

== How To Add The Macro ==
1. Select the 'Developer' tab on the ribbon and select 'Insert' -> 'Button'
2. Place the button at a desired location on the spreadsheet. Note: Button should be outside of the data source range (A1:BZ5000) such as column CA. 
3. Select 'New'
4. In the workbook code area that appears add 'ExcelClone.Show'. Should look like:

 		Sub Button1_Click()
		ExcelClone.Show
		End Sub
5. Import the file 'ExcelClone.frm'
6. Right click the button on the Spreadsheet and select 'Edit Text'
7. Rename the button accordingly. Example: 'Custom Report'


== How To Use ==
1. Select the "Custom Report" button at the end of the first row, near the "CA" column. 
2. The user interface (UI) appears allowing you to edit certain fields:
	2.1 Report Name: Name for the output .xlsx report. The report name input will be preceeded by 'COPY_ONLY_' and followed by a timestamp. Example: COPY_ONLY_MyReportName_12-14-16-10-15.xlsx
	2.2 Source Sheet Name: Name of the sheet tab with the source data. Default name is 'Sheet1'
	2.3 Select Columns: Check boxes for selecting each column to be reported or use 'Select All' to report all columns
3. Once columns are selected click 'Submit' 
4. A new file will generate, closing the UI. 
5. The new file will contain
	5.1 Single cell in A1 that is locked and password protected to identify the sheet as a Copy
	5.2 The custom report in range A2:BZ5001 with only the selected columns 


== Frequently Asked Questions ==

= Where is the 'CustomReport.frx' file? =

The file was provided as part of the bundle with this readme file. Look for hte file at the location where you extracted the file. 

= Where does the new file save? =

By default the new file will save on your desktop. 

= What happens with hidden rows and columns? =

The output copies ALL rows 1 through 5000 of the selected columns even if the column is hidden. 



= Troubleshooting =

If things don't work when you set the macro check the following:

1.  Is the button assigned the correct macro. 
2.  Does the button macro call the correct form name? ('CustomReport.Show')


== Licenses ==
Copyright (c) 2016 Jonathan Vargas @ https://github.com/sayuga

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

== Contact Information ==

You may contact the developer at jvargas@fletchermartincorp.com or through https://github.com/sayuga
