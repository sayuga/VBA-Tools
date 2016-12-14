VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExcelClone 
   Caption         =   "ExcelClone"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7815
   OleObjectBlob   =   "ExcelClone.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExcelClone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub CheckBoxAP_Click()

End Sub

Private Sub CloseButton_Click()
Unload Me
End Sub

Private Sub UserForm_Initialize()

reportName.Value = ""
sheetSourceName.Value = "Sheet1"

'COLUMNS A - Z
CheckBoxA.Value = False
CheckBoxB.Value = False
CheckBoxC.Value = False
CheckBoxD.Value = False
CheckBoxE.Value = False
CheckBoxF.Value = False
CheckBoxG.Value = False
CheckBoxH.Value = False
CheckBoxI.Value = False
CheckBoxJ.Value = False
CheckBoxK.Value = False
CheckBoxL.Value = False
CheckBoxM.Value = False
CheckBoxN.Value = False
CheckBoxO.Value = False
CheckBoxP.Value = False
CheckBoxQ.Value = False
CheckBoxR.Value = False
CheckBoxS.Value = False
CheckBoxT.Value = False
CheckBoxU.Value = False
CheckBoxV.Value = False
CheckBoxW.Value = False
CheckBoxX.Value = False
CheckBoxY.Value = False
CheckBoxZ.Value = False

'COLUMNS AA - AZ
CheckBoxAA.Value = False
CheckBoxAB.Value = False
CheckBoxAC.Value = False
CheckBoxAD.Value = False
CheckBoxAE.Value = False
CheckBoxAF.Value = False
CheckBoxAG.Value = False
CheckBoxAH.Value = False
CheckBoxAI.Value = False
CheckBoxAJ.Value = False
CheckBoxAK.Value = False
CheckBoxAL.Value = False
CheckBoxAM.Value = False
CheckBoxAN.Value = False
CheckBoxAO.Value = False
CheckBoxAP.Value = False
CheckBoxAQ.Value = False
CheckBoxAR.Value = False
CheckBoxAS.Value = False
CheckBoxAT.Value = False
CheckBoxAU.Value = False
CheckBoxAV.Value = False
CheckBoxAW.Value = False
CheckBoxAX.Value = False
CheckBoxAY.Value = False
CheckBoxAZ.Value = False

'COLUMNS BA - BZ
CheckBoxBA.Value = False
CheckBoxBB.Value = False
CheckBoxBC.Value = False
CheckBoxBD.Value = False
CheckBoxBE.Value = False
CheckBoxBF.Value = False
CheckBoxBG.Value = False
CheckBoxBH.Value = False
CheckBoxBI.Value = False
CheckBoxBJ.Value = False
CheckBoxBK.Value = False
CheckBoxBL.Value = False
CheckBoxBM.Value = False
CheckBoxBN.Value = False
CheckBoxBO.Value = False
CheckBoxBP.Value = False
CheckBoxBQ.Value = False
CheckBoxBR.Value = False
CheckBoxBS.Value = False
CheckBoxBT.Value = False
CheckBoxBU.Value = False
CheckBoxBV.Value = False
CheckBoxBW.Value = False
CheckBoxBX.Value = False
CheckBoxBY.Value = False
CheckBoxBZ.Value = False

CheckBoxSelectAll.Value = False

reportName.SetFocus
End Sub
Private Sub reportName_Change()
reportName.Text = reportName.Value
End Sub
Private Sub sheetSourceName_Change()
sheetSourceName.Text = sheetSourceName.Value
End Sub

Private Sub SubmitButton_Click()
  Dim sheetName As String
    Dim dt As String
    Dim waterMark As String
    Dim repName As String
    Dim wbNew As Worksheet
    Dim wbOrig As Worksheet
    Dim i As Integer
    
    i = 1
    sheetName = sheetSourceName.Value
    dt = "_" & Format(CStr(Now), "mm_dd_yy_hh_mm")
    waterMark = "COPY_ONLY_"
    repName = reportName.Value
    
    Set wbOrig = ThisWorkbook.Worksheets(sheetName)
    Set wbNew = Workbooks.Add.Worksheets(sheetName)
        
    'COLUMNS A THROUGH Z
    If CheckBoxA.Value = True Then
    wbOrig.Range("A1:A5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxB.Value = True Then
    wbOrig.Range("B1:B5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxC.Value = True Then
    wbOrig.Range("C1:C5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxD.Value = True Then
    wbOrig.Range("D1:D5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxE.Value = True Then
    wbOrig.Range("E1:E5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxF.Value = True Then
    wbOrig.Range("F1:F5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxG.Value = True Then
    wbOrig.Range("G1:G5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxH.Value = True Then
    wbOrig.Range("H1:H5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxI.Value = True Then
    wbOrig.Range("I1:I5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxJ.Value = True Then
    wbOrig.Range("J1:J5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxK.Value = True Then
    wbOrig.Range("K1:K5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxL.Value = True Then
    wbOrig.Range("L1:L5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxM.Value = True Then
    wbOrig.Range("M1:M5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxN.Value = True Then
    wbOrig.Range("N1:N5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxO.Value = True Then
    wbOrig.Range("O1:O5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxP.Value = True Then
    wbOrig.Range("P1:P5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxQ.Value = True Then
    wbOrig.Range("Q1:Q5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxR.Value = True Then
    wbOrig.Range("R1:R5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxS.Value = True Then
    wbOrig.Range("S1:S5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxT.Value = True Then
    wbOrig.Range("T1:T5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxU.Value = True Then
    wbOrig.Range("U1:U5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxV.Value = True Then
    wbOrig.Range("V1:V5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxW.Value = True Then
    wbOrig.Range("W1:W5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxX.Value = True Then
    wbOrig.Range("X1:X5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxY.Value = True Then
    wbOrig.Range("Y1:Y5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxZ.Value = True Then
    wbOrig.Range("Z1:Z5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
        
    'COLUMNS AA THROUGH AZ
      If CheckBoxAA.Value = True Then
    wbOrig.Range("AA1:AA5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAB.Value = True Then
    wbOrig.Range("AB1:AB5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAC.Value = True Then
    wbOrig.Range("AC1:AC5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAD.Value = True Then
    wbOrig.Range("AD1:AD5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAE.Value = True Then
    wbOrig.Range("AE1:AE5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAF.Value = True Then
    wbOrig.Range("AF1:AF5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAG.Value = True Then
    wbOrig.Range("AG1:AG5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAH.Value = True Then
    wbOrig.Range("AH1:AH5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAI.Value = True Then
    wbOrig.Range("AI1:AI5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAJ.Value = True Then
    wbOrig.Range("AJ1:AJ5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAK.Value = True Then
    wbOrig.Range("AK1:AK5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAL.Value = True Then
    wbOrig.Range("AL1:AL5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAM.Value = True Then
    wbOrig.Range("AM1:AM5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAN.Value = True Then
    wbOrig.Range("AN1:AN5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAO.Value = True Then
    wbOrig.Range("AO1:AO5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAP.Value = True Then
    wbOrig.Range("AP1:AP5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAQ.Value = True Then
    wbOrig.Range("AQ1:AQ5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAR.Value = True Then
    wbOrig.Range("AR1:AR5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAS.Value = True Then
    wbOrig.Range("AS1:AS5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAT.Value = True Then
    wbOrig.Range("AT1:AT5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAU.Value = True Then
    wbOrig.Range("AU1:AU5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAV.Value = True Then
    wbOrig.Range("AV1:AV5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAW.Value = True Then
    wbOrig.Range("AW1:AW5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAX.Value = True Then
    wbOrig.Range("AX1:AX5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAY.Value = True Then
    wbOrig.Range("AY1:AY5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxAZ.Value = True Then
    wbOrig.Range("AZ1:AZ5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    
    ' COlUMNS BA THROUGH BZ
      If CheckBoxBA.Value = True Then
    wbOrig.Range("BA1:BA5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBB.Value = True Then
    wbOrig.Range("BB1:BB5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBC.Value = True Then
    wbOrig.Range("BC1:BC5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBD.Value = True Then
    wbOrig.Range("BD1:BD5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBE.Value = True Then
    wbOrig.Range("BE1:BE5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBF.Value = True Then
    wbOrig.Range("BF1:BF5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBG.Value = True Then
    wbOrig.Range("BG1:BG5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBH.Value = True Then
    wbOrig.Range("BH1:BH5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBI.Value = True Then
    wbOrig.Range("BI1:BI5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBJ.Value = True Then
    wbOrig.Range("BJ1:BJ5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBK.Value = True Then
    wbOrig.Range("BK1:BK5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBL.Value = True Then
    wbOrig.Range("BL1:BL5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBM.Value = True Then
    wbOrig.Range("BM1:BM5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBN.Value = True Then
    wbOrig.Range("BN1:BN5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBO.Value = True Then
    wbOrig.Range("BO1:BO5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBP.Value = True Then
    wbOrig.Range("BP1:BP5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBQ.Value = True Then
    wbOrig.Range("BQ1:BQ5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBR.Value = True Then
    wbOrig.Range("BR1:BR5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBS.Value = True Then
    wbOrig.Range("BS1:BS5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBT.Value = True Then
    wbOrig.Range("BT1:BT5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBU.Value = True Then
    wbOrig.Range("BU1:BU5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBV.Value = True Then
    wbOrig.Range("BV1:BV5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBW.Value = True Then
    wbOrig.Range("BW1:BW5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBX.Value = True Then
    wbOrig.Range("BX1:BX5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBY.Value = True Then
    wbOrig.Range("BY1:BY5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    If CheckBoxBZ.Value = True Then
    wbOrig.Range("BZ1:BZ5000").Copy wbNew.Cells(2, i)
    i = i + 1
    End If
    
    
    'SELECT ALL
    If CheckBoxSelectAll.Value = True Then
    wbOrig.Range("A1:BZ5000").Copy wbNew.Cells(2, i)
    End If
        
    wbNew.Range("A1").Value = waterMark
    ActiveCell.Offset(0, 1).Activate
    ActiveCell.Font.Bold = True
    wbNew.Columns.AutoFit
    wbNew.Range("B1:BZ500").Locked = False
    wbNew.Range("A2:A500").Locked = False
    wbNew.Protect Password:="Spinelli36"
    
wbNam = waterMark & reportName & dt
    ActiveWorkbook.SaveAs Filename:=wbNam
   
End Sub
