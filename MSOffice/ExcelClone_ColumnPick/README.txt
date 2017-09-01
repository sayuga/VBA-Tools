
Developer: Jonathan Vargas 
Version: 1.0.0
Revision Date: 09-01-2017
Requires: Macro-Enabled Excel File
Copyright: MIT/X11

Description: 
Macro function that allows a clone of columns from an active worksheet into a new workbook.
Optional columns definition allows to define the columns to be copied over thus allowing custom reports with a single click. 

How to use: 
- Open Excel and go to developer tab. 
- Select 'View Code'
- Add the function to the 'ThisWorkbook' project file. 
- Call the funciton using 'customRangeReporter'
  -- Parameters
    -+ hasList: Type Boolean. User defined to use or not a column list
    -+ ReportName: Type String. Name for report. *Report Names will be followed by the watermark when saving the file
    -+ wmPassword: Type String. Password locking the Watermarked cell at "A1"
    -+ colList: Type Variant. String Array with a list of the columns to be copied
- Go back to the excel sheet and assign the your macro to a button or custom ribbon. 
- Click the button/ribbon option to test. 

