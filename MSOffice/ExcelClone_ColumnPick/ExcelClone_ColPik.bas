Sub FullCopyCustomReport_Click()    
    Dim ReportName As String
    Dim watermarkPassword As String
    Dim hasList As Boolean
    
    hasList = False
    ReportName = "FullCopy_"
    watermarkPassword = "ABC123"
    
    customRangeReport hasList, ReportName, watermarkPassword
End Sub

Sub CustomReport_Click()
    Dim ColumnsList
    Dim ReportName As String
    Dim watermarkPassword As String
    Dim hasList As Boolean
    
    hasList = True
    ColumnsList = Array("A", "B", "J", "K", "M", "O", "P", "q", "R", "U")
    ReportName = "SelectedColumn_"
    watermarkPassword = "ABC123"
    
    customRangeReport hasList, ReportName, watermarkPassword, ColumnsList
End Sub


Function customRangeReport(hasList As Boolean, ReportName As String, wmPassword As String, Optional colList As Variant)
    
    'colList is a array of the column letters you want in your report        
    'rName is the report name    
    
    Dim dt, waterMark, sheetName As String
    Dim wbNew, wbOrig As Worksheet
    Dim c1, c2 As Range    
    Dim I As Integer
    Dim fName As String
    Dim desktopFolderPath As String
    I = 1
    waterMark = "COPY_ONLY_Report_Generated_On_" & Format(CStr(Now), "mm_dd_yy_hh_mm")
        
    With ActiveSheet
        'get last col with data
        lCol = .Cells(1, Columns.Count).End(xlToLeft).Column
        'get last row with data
        lRow = .Cells(Rows.Count, 1).End(xlUp).Row
        'get sheet name
        sheetName = .Name
    End With
    
    Set wbOrig = ThisWorkbook.Worksheets(sheetName)
    Set c1 = wbOrig.Range(Cells(1, 1), Cells(lRow, lCol))
    
    Set wbNew = Workbooks.Add.Sheets(1)
    With wbNew
        'Create the Watermarked cell that will be 
        'Locked at the end. 
        .Name = sheetName
        With .Cells(1, 1)
            .Value = waterMark
            .Font.Bold = True
        End With
        
        'Checks if a columns list was 
        If hasList = True Then
            'If there is a list, adds each column to the new
            'copy in the order it was provided. 
            For Each Item In colList
                x = Item
                q = x & "1:" & x & lRow
                Set c1 = wbOrig.Range(q)
                Set c2 = wbNew.Cells(2, I)
                c1.Copy c2
                I = I + 1
            Next
        ' If no list provided, copies all the columns
        Else           
            Set c2 = wbNew.Range(Cells(2, 1), Cells(lRow + 1, lCol))
            c1.Copy c2
        End If
        
        'Remove locking from all active cells except for "A1"
        'And resize the columns to make them fit better
        Set c2 = .Range(Cells(2, 1), Cells(lRow + 1, lCol))
        c2.Locked = False        
        Set c2 = .Range(Cells(1, 2), Cells(1, lCol))
        c2.Locked = False        
        Set c2 = .Range(Cells(2, 1), Cells(2, I))
        c2.WrapText = False
        c2.Locked = False        
        .Columns.AutoFit
        
        'Add a password for locking "A1" a.k.a. the Watermark        
        .Protect Password:= wmPassword
        
        'Save File to Users Desktop
        desktopFolderPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
        fName = desktopFolderPath & ReportName & waterMark & ".xlsx" 'I add a time stamp for auditing process. 
        .SaveAs fName
    End With
End Function
