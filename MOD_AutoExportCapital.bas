Attribute VB_Name = "MOD_AutoExportCapital"
Option Explicit
'Purpose: To house various subs for supporting the AUTOMATIC CREATION OF PROSYSTEM CAPITAL GAINS EXPORT feature of the KTC Tax Excel Add-in

Private Sub SUB_Auto_Export_Capital(Optional Control As IRibbonControl)
    'Author: Edward J. Sullivan edward.sullivan@kinshiptrustco.com
    'Purpose: To automatically export capital items from the gain-loss sheet to a prosystem export sheet.

    If pfBLN_ActiveSheetExists(BLN_DispMsg:=True) = False Then Exit Sub
    If pfBLN_SheetExists(STR_SheetName:="GAIN - LOSS", BLN_DispMsg:=True) = False Then Exit Sub

    pSTR_SUBFCN_Name = "SUB_Auto_Export_Capital"
    On Error GoTo ErrorHandler

    pSUB_PerformanceOptions BLN_TurnOn:=True

    Set pWB_User = ActiveWorkbook
    SUB_GainExportExists

    Dim LNG_NextEmptyRow As Long
    Dim WS_GainLoss As Worksheet
    Set WS_GainLoss = pWB_User.Sheets("GAIN - LOSS")
    Dim WS_CapGainExport As Worksheet
    Set WS_CapGainExport = pWB_User.Sheets("GAIN EXPORT")

    WS_CapGainExport.Range("A7:AC500").ClearContents    'Clear sheet first.

    For pINT_Target = LBound(fARR_CapGainSectionName) To UBound(fARR_CapGainSectionName)

        For Each pRNG_Target In fARR_CapGainTargetRange(WS_GainLoss)(pINT_Target)

            If pRNG_Target.Value <> 0 And _
               pRNG_Target.Offset(0, 1).Value <> "P" Then    ' Skip passive entries that will be marked with a 'P'

                SUB_TransferCapitalData LNG_NextEmptyRow, WS_CapGainExport

            End If
        Next
    Next

    SUB_SelectExportRange LNG_NextEmptyRow, WS_CapGainExport
    Debug.Print "Capital Gains Export has been created."

    pSUB_PerformanceOptions BLN_TurnOn:=False

    On Error GoTo 0
    Exit Sub
ErrorHandler:
    pSUB_GlobalErrorHandler

End Sub

Private Sub SUB_GainExportExists()
    'Author: Edward J. Sullivan edward.sullivan@kinshiptrustco.com
    'Purpose: Check to see if "GAIN EXPORT" worksheet exists, if not, add it.

    pSTR_SUBFCN_Name = "SUB_GainExportExists"

    If pfBLN_SheetExists(STR_SheetName:="GAIN EXPORT", BLN_DispMsg:=False) = False Then

        Dim WB_CapGainExport As Workbook
        Set WB_CapGainExport = Workbooks.Open(Filename:="J:\TAX\KTC Tax Excel Add-in\Template_1040\Capital_Gain_Export.xlsx", ReadOnly:=True)
        WB_CapGainExport.Sheets("GAIN EXPORT").Move before:=pWB_User.Sheets("GAIN - LOSS")    'Don't need to close since we moved the only tab out.
        Debug.Print "Sheet 'GAIN EXPORT' Added."

    End If

End Sub

Private Function fARR_CapGainSectionName() As Variant
    'Author: Edward J. Sullivan edward.sullivan@kinshiptrustco.com
    'Purpose: To establish the array named
    'Note: All the arrays used in this module need to be ordered the same.

    pSTR_SUBFCN_Name = "fARR_CapGainSectionName"

    Dim ARR_CapGainSectionName(11) As String

    ARR_CapGainSectionName(0) = "Reg_ST_BoxA"
    ARR_CapGainSectionName(1) = "Reg_ST_BoxB"
    ARR_CapGainSectionName(2) = "Reg_ST_BoxC"
    ARR_CapGainSectionName(3) = "Reg_LT_BoxA"
    ARR_CapGainSectionName(4) = "Reg_LT_BoxB"
    ARR_CapGainSectionName(5) = "Reg_LT_BoxC"
    ARR_CapGainSectionName(6) = "AMT_ST_BoxA"
    ARR_CapGainSectionName(7) = "AMT_ST_BoxB"
    ARR_CapGainSectionName(8) = "AMT_ST_BoxC"
    ARR_CapGainSectionName(9) = "AMT_LT_BoxA"
    ARR_CapGainSectionName(10) = "AMT_LT_BoxB"
    ARR_CapGainSectionName(11) = "AMT_LT_BoxC"

    fARR_CapGainSectionName = ARR_CapGainSectionName

End Function

Private Function fARR_CapGainTargetRange(WS_GainLoss As Worksheet) As Variant
    'Author: Edward J. Sullivan edward.sullivan@kinshiptrustco.com
    'Purpose: To establish the array named
    'Note: All the arrays used in this module need to be ordered the same.

    pSTR_SUBFCN_Name = "fARR_CapGainTargetRange"

    Dim ARR_CapGainTargetRange(11) As Range

    With WS_GainLoss
        Set ARR_CapGainTargetRange(0) = .Range("B16:B65,B69:B118")
        Set ARR_CapGainTargetRange(1) = .Range("B124:B173,B177:B173")
        Set ARR_CapGainTargetRange(2) = .Range("B232:B281,B285:B334")
        Set ARR_CapGainTargetRange(3) = .Range("B351:B399,B404:B453")
        Set ARR_CapGainTargetRange(4) = .Range("B459:B508,B512:B561")
        Set ARR_CapGainTargetRange(5) = .Range("B567:B616,B620:B669")
        Set ARR_CapGainTargetRange(6) = .Range("B687:B736,B740:B789")
        Set ARR_CapGainTargetRange(7) = .Range("B795:B844,B848:B898")
        Set ARR_CapGainTargetRange(8) = .Range("B903:B952,B956:B1005")
        Set ARR_CapGainTargetRange(9) = .Range("B1022:B1071,B1075:B1124")
        Set ARR_CapGainTargetRange(10) = .Range("B1130:B1179,B1183:B1232")
        Set ARR_CapGainTargetRange(11) = .Range("B1238:B1287,B1291:B1340")
    End With

    fARR_CapGainTargetRange = ARR_CapGainTargetRange

End Function

Private Function fARR_CapGainTermCode() As Variant
    'Author: Edward J. Sullivan edward.sullivan@kinshiptrustco.com
    'Purpose: To establish the array named
    'Note: All the arrays used in this module need to be ordered the same.

    pSTR_SUBFCN_Name = "fARR_CapGainTermCode"

    Dim ARR_CapGainTermCode(11) As String

    ARR_CapGainTermCode(0) = "S"
    ARR_CapGainTermCode(1) = "S"
    ARR_CapGainTermCode(2) = "S"
    ARR_CapGainTermCode(3) = "L"
    ARR_CapGainTermCode(4) = "L"
    ARR_CapGainTermCode(5) = "L"
    ARR_CapGainTermCode(6) = "S"
    ARR_CapGainTermCode(7) = "S"
    ARR_CapGainTermCode(8) = "S"
    ARR_CapGainTermCode(9) = "L"
    ARR_CapGainTermCode(10) = "L"
    ARR_CapGainTermCode(11) = "L"

    fARR_CapGainTermCode = ARR_CapGainTermCode

End Function

Private Function fARR_CapGain1099BCode() As Variant
    'Author: Edward J. Sullivan edward.sullivan@kinshiptrustco.com
    'Purpose: To establish the array named
    'Note: All the arrays used in this module need to be ordered the same.

    pSTR_SUBFCN_Name = "fARR_CapGain1099BCode"

    Dim ARR_CapGain1099BCode(11) As String

    ARR_CapGain1099BCode(0) = "A"
    ARR_CapGain1099BCode(1) = "B"
    ARR_CapGain1099BCode(2) = "C"
    ARR_CapGain1099BCode(3) = "A"
    ARR_CapGain1099BCode(4) = "B"
    ARR_CapGain1099BCode(5) = "C"
    ARR_CapGain1099BCode(6) = "A"
    ARR_CapGain1099BCode(7) = "B"
    ARR_CapGain1099BCode(8) = "C"
    ARR_CapGain1099BCode(9) = "A"
    ARR_CapGain1099BCode(10) = "B"
    ARR_CapGain1099BCode(11) = "C"

    fARR_CapGain1099BCode = ARR_CapGain1099BCode

End Function

Private Function fARR_CapGainAMTCode() As Variant
    'Author: Edward J. Sullivan edward.sullivan@kinshiptrustco.com
    'Purpose: To establish the array named
    'Note: All the arrays used in this module need to be ordered the same.

    pSTR_SUBFCN_Name = "fARR_CapGainAMTCode"

    Dim ARR_CapGainAMTCode(11) As String

    ARR_CapGainAMTCode(0) = "1"
    ARR_CapGainAMTCode(1) = "1"
    ARR_CapGainAMTCode(2) = "1"
    ARR_CapGainAMTCode(3) = "1"
    ARR_CapGainAMTCode(4) = "1"
    ARR_CapGainAMTCode(5) = "1"
    ARR_CapGainAMTCode(6) = "2"
    ARR_CapGainAMTCode(7) = "2"
    ARR_CapGainAMTCode(8) = "2"
    ARR_CapGainAMTCode(9) = "2"
    ARR_CapGainAMTCode(10) = "2"
    ARR_CapGainAMTCode(11) = "2"

    fARR_CapGainAMTCode = ARR_CapGainAMTCode

End Function

Private Sub SUB_TransferCapitalData(LNG_NextEmptyRow As Long, WS_CapGainExport As Worksheet)
    'Author: Edward J. Sullivan edward.sullivan@kinshiptrustco.com
    'Purpose: To transfer data between gain-loss worksheet and the prosystem export sheet.

    pSTR_SUBFCN_Name = "SUB_TransferCapitalData"

    Dim STR_Term_Code As String
    Dim STR_1099B_Code As String
    Dim STR_AMT_Code As String

    STR_Term_Code = fARR_CapGainTermCode(pINT_Target)
    STR_1099B_Code = fARR_CapGain1099BCode(pINT_Target)
    STR_AMT_Code = fARR_CapGainAMTCode(pINT_Target)

    Dim STR_Description_Col As String
    Dim STR_Sales_Price_Col As String
    Dim STR_Cost_Basis_Col As String
    Dim STR_Date_Acquired_Col As String
    Dim STR_Date_Sold_Col As String
    Dim STR_Term_Code_Col As String
    Dim STR_1099B_Code_Col As String
    Dim STR_8949_Code_Col As String
    Dim STR_Adjustment_Col As String
    Dim STR_AMT_Code_Col As String

    STR_Description_Col = "A"
    STR_Sales_Price_Col = "C"
    STR_Cost_Basis_Col = "D"
    STR_Date_Acquired_Col = "F"
    STR_Date_Sold_Col = "G"
    STR_Term_Code_Col = "H"
    STR_1099B_Code_Col = "I"
    STR_8949_Code_Col = "K"
    STR_Adjustment_Col = "Q"
    STR_AMT_Code_Col = "S"

    With WS_CapGainExport
        LNG_NextEmptyRow = .Cells(Rows.Count, STR_Description_Col).End(xlUp).Row + 1    ' Find next empty row
        .Cells(LNG_NextEmptyRow, STR_Description_Col).Value = pRNG_Target.Value
        .Cells(LNG_NextEmptyRow, STR_Sales_Price_Col).Value = pRNG_Target.Offset(0, 5).Value
        .Cells(LNG_NextEmptyRow, STR_Cost_Basis_Col).Value = pRNG_Target.Offset(0, 6).Value
        .Cells(LNG_NextEmptyRow, STR_Date_Acquired_Col).Value = pRNG_Target.Offset(0, 3).Value
        .Cells(LNG_NextEmptyRow, STR_Date_Sold_Col).Value = pRNG_Target.Offset(0, 4).Value
        .Cells(LNG_NextEmptyRow, STR_Term_Code_Col).Value = STR_Term_Code
        .Cells(LNG_NextEmptyRow, STR_1099B_Code_Col).Value = STR_1099B_Code
        .Cells(LNG_NextEmptyRow, STR_8949_Code_Col).Value = pRNG_Target.Offset(0, 7).Value
        .Cells(LNG_NextEmptyRow, STR_Adjustment_Col).Value = pRNG_Target.Offset(0, 8).Value
        .Cells(LNG_NextEmptyRow, STR_AMT_Code_Col).Value = STR_AMT_Code
    End With

End Sub

Private Sub SUB_SelectExportRange(LNG_NextEmptyRow As Long, WS_CapGainExport As Worksheet)
    'Author: Edward J. Sullivan edward.sullivan@kinshiptrustco.com
    'Purpose: To select the export range for the user to copy and paste.

    pSTR_SUBFCN_Name = "SUB_SelectExportRange"

    With WS_CapGainExport
        .Visible = xlSheetVisible    ' Make sure sheet is not hidden.
        .Activate    'Activate gain export sheet so that we can select the export range.
        LNG_NextEmptyRow = .Cells(Rows.Count, "A").End(xlUp).Row
        .Range(Cells(7, "A"), Cells(LNG_NextEmptyRow, "AC")).Select    'Select range for user.
    End With

End Sub
