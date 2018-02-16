Attribute VB_Name = "MOD_AutoImportK1"
Option Explicit
'Author: Edward Sullivan edwardjohnsullivan@gmail.com
'Purpose: To house various procedures to support the K-1 Auto Import Feature.


Private Sub SUB_Auto_Import_K1_Partner(Optional Control As IRibbonControl)
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To automatically import K-1 values from network stored K-1 Outputs to a user's Tax K-1 Summary

    If pfBLN_ActiveSheetExists(BLN_DispMsg:=True) = False Then Exit Sub
    If pfBLN_SheetExists(STR_SheetName:="K-1 SUMMARY", BLN_DispMsg:=True) = False Then Exit Sub
    If pfBLN_IsWpUpToDate(STR_WpName:="1040", BLN_DispUpdateMsg:=False, BLN_DispErrMsg:=True) = False Then Exit Sub
    If fBLN_UserAcceptsDisclaimer = False Then Exit Sub

    pSTR_SUBFCN_Name = "SUB_Auto_Import_K1_Partner"
    On Error GoTo ErrorHandler

    pSUB_PerformanceOptions BLN_TurnOn:=True
    Dim SGL_StartTime As Single
    Dim SGL_TimeElapsed As Single
    pSUB_TimeElapsed SGL_StartTime, SGL_TimeElapsed, BLN_StartTimer:=True
    
    pSUB_SetDynamicStatusMsg STR_Msg:="Saving Current Tax 1040 Workpaper."

    Dim WB_1040 As Workbook
    Set WB_1040 = ActiveWorkbook
    WB_1040.Save

    Dim WS_1040K1Summ As Worksheet
    Set WS_1040K1Summ = WB_1040.Sheets("K-1 SUMMARY")
    WS_1040K1Summ.Activate    ' Not required, just for presentation at the end.

    Dim RNG_1040eIDLookup As Range
    Set RNG_1040eIDLookup = WS_1040K1Summ.Range("I1:I878")
    
    pSUB_SetDynamicStatusMsg STR_Msg:="Opening Entity ID Workbook."

    Dim WB_eID As Workbook
    Set WB_eID = Workbooks.Open(Filename:="J:\TAX\Tax Excel Add-in\External_Variables\ENTITY_ID.xlsx", ReadOnly:=True)
    Dim WS_eID As Worksheet
    Set WS_eID = WB_eID.Sheets("ENTITY_ID")
    Dim RNG_eIDTarget As Range
    Dim RNG_eIDLookup As Range
    Dim RNG_eIDPathLookup As Range

    pINT_Counter = 0

    Dim BLN_eIDErrExists As Boolean
    BLN_eIDErrExists = False
    Dim BLN_PathErrExists As Boolean
    BLN_PathErrExists = False
    Dim BLN_InvalidPathExists As Boolean
    BLN_InvalidPathExists = False
    Dim BLN_pIDErrExists As Boolean
    BLN_pIDErrExists = False
    Dim BLN_TaxIncErrExists As Boolean
    BLN_TaxIncErrExists = False

    Dim OBJ_ExpLog As Object
    Dim STR_ExpLogPath As String
    SUB_OpenExportLog WB_1040, STR_ExpLogPath, OBJ_ExpLog
    
    pSUB_SetDynamicStatusMsg STR_Msg:="Starting Import Process."

    For Each RNG_eIDTarget In RNG_1040eIDLookup

        Dim STR_1040eID As String

        SUB_SeteIDVars RNG_eIDTarget, STR_1040eID, WS_1040K1Summ, RNG_1040eIDLookup

        If fBLN_eIDExists(STR_1040eID) Then

            If fBLN_IseIDValid(BLN_eIDErrExists, STR_1040eID, WS_1040K1Summ, _
                               RNG_1040eIDLookup, OBJ_ExpLog) Then

                Dim WB_K1Out As Workbook
                Dim STR_1040pID As String
                Dim RNG_K1OutpIDLookup As Range
                Dim STR_1040pName As String
                Dim STR_K1OutPath As String
                STR_K1OutPath = fSTR_K1OutPath(STR_1040eID, WS_eID, RNG_eIDLookup, RNG_eIDPathLookup)

                If fBLN_K1OutPathExists(BLN_PathErrExists, STR_1040eID, STR_K1OutPath, OBJ_ExpLog) Then

                    If fBLN_K1OutPathIsValid(BLN_InvalidPathExists, STR_1040eID, STR_K1OutPath, OBJ_ExpLog) Then

                        Dim WS_K1Out As Worksheet

                        SUB_SetpIDVars RNG_eIDTarget, STR_1040pID, WB_K1Out, _
                                       WS_K1Out, STR_K1OutPath, RNG_K1OutpIDLookup, _
                                       STR_1040pName

                        If fBLN_pIDIsValid(BLN_pIDErrExists, STR_1040pID, RNG_K1OutpIDLookup, _
                                           STR_1040pName, OBJ_ExpLog) Then

                            Dim RNG_1040pIDLookup As Range
                            Dim RNG_K1OuteIDLookup As Range
                            Dim LNG_1040ImportRow As Long
                            Dim LNG_K1OutImportRow As Long
                            Dim RNG_1040Date As Range
                            Dim RNG_K1OutDate As Range

                            SUB_SetDateVars STR_1040pID, RNG_1040pIDLookup, RNG_K1OutpIDLookup, _
                                            LNG_1040ImportRow, WS_1040K1Summ, RNG_1040Date, _
                                            LNG_K1OutImportRow, WS_K1Out, RNG_K1OutDate

                            If fBLN_DateIsNotCurrent(STR_1040pID, RNG_1040Date, RNG_K1OutDate, _
                                                     STR_1040pName, OBJ_ExpLog) Then

                                Dim RNG_1040Import As Range
                                Dim RNG_K1OutExport As Range
                                Dim RNG_1040EstorAct As Range
                                Dim RNG_K1OutEstorAct As Range

                                SUB_SetImportVars RNG_1040Import, WS_1040K1Summ, LNG_1040ImportRow, _
                                                  RNG_K1OutExport, WS_K1Out, LNG_K1OutImportRow, _
                                                  RNG_1040EstorAct, RNG_K1OutEstorAct

                                SUB_PerformImport RNG_1040Import, RNG_K1OutExport, RNG_1040Date, _
                                                  RNG_K1OutDate, RNG_1040EstorAct, RNG_K1OutEstorAct, _
                                                  STR_1040pName

                                Dim RNG_1040TaxInc As Range
                                Dim RNG_K1OutTaxInc As Range

                                SUB_SetTaxIncVars WS_1040K1Summ, RNG_1040TaxInc, LNG_1040ImportRow, _
                                                  WS_K1Out, RNG_K1OutTaxInc, LNG_K1OutImportRow

                                SUB_CheckTaxInc BLN_TaxIncErrExists, WS_1040K1Summ, RNG_1040TaxInc, _
                                                WS_K1Out, RNG_K1OutTaxInc, STR_1040pName, _
                                                RNG_1040Import, STR_1040pID, OBJ_ExpLog

                            End If
                        End If

                        With WB_K1Out
                            .Saved = True
                            .Close    'We are closing out the target workbook every loop until we can find a better way.
                                      'There is a performance cost here.
                        End With

                    End If
                End If
            End If
        End If
    Next

    OBJ_ExpLog.Close

    With WB_eID
        .Saved = True
        .Close
    End With
    
    pSUB_PerformanceOptions BLN_TurnOn:=False

    pSUB_TimeElapsed SGL_StartTime, SGL_TimeElapsed, BLN_StartTimer:=False
    SUB_CompletionMsg BLN_pIDErrExists, BLN_eIDErrExists, BLN_TaxIncErrExists, _
                      BLN_PathErrExists, BLN_InvalidPathExists, STR_ExpLogPath, _
                      SGL_TimeElapsed
    
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    OBJ_ExpLog.Close
    pSUB_GlobalErrorHandler

End Sub

Private Function fSTR_K1Col(STR_ColName As String) As String
    ' Author: Edward Sullivan edwardjohnsullivan@gmail.com
    ' Purpose: To return the column name as needed.

    pSTR_SUBFCN_Name = "fSTR_K1Col"

    Select Case STR_ColName

        Case "FirstImportCol"
            fSTR_K1Col = "AC"

        Case "LastImportCol"
            fSTR_K1Col = "GA"

        Case "ESTorACTCol"
            fSTR_K1Col = "E"

        Case "TaxableIncomeCol"
            fSTR_K1Col = "GF"

        Case "DateCol"
            fSTR_K1Col = "K"

    End Select

End Function

Private Function fBLN_UserAcceptsDisclaimer() As Boolean
    ' Author: Edward Sullivan edwardjohnsullivan@gmail.com
    ' Purpose: To determine if the user accepted the STR_Initial disclaimer message.

    pSTR_SUBFCN_Name = "fBLN_UserAcceptsDisclaimer"
    fBLN_UserAcceptsDisclaimer = False

    If MsgBox("Do you want to import K-1 values from available Prepared K-1 OUTPUT sheets?" & vbNewLine & vbNewLine & _
              "This action will replace ALL values in the rows of the supported entity lines of your Tax 1040 K-1 Summary." & vbNewLine & vbNewLine & _
              "A text file export log will be generated in the same directory as this Workpaper Set.", _
              vbYesNo, "Automatic Import of K-1 Values") = vbYes Then

        fBLN_UserAcceptsDisclaimer = True

    End If

    #If CON_BLN_DebugModeActivated Then
        Debug.Print pSTR_SUBFCN_Name & " = " & fBLN_UserAcceptsDisclaimer
    #End If

End Function

Private Sub SUB_OpenExportLog(WB_1040 As Workbook, STR_ExpLogPath As String, OBJ_ExpLog As Object)
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To create/open the text file export log for use.

    pSTR_SUBFCN_Name = "SUB_OpenExportLog"
    pSUB_SetDynamicStatusMsg STR_Msg:="Creating Export Log."

    Dim OBJ_FSO As Object
    Set OBJ_FSO = CreateObject("Scripting.FileSystemObject")
    STR_ExpLogPath = WB_1040.Path & "\" & WB_1040.Name & "-EXPORT LOG.txt"
    Set OBJ_ExpLog = OBJ_FSO.CreateTextFile(STR_ExpLogPath, fsoForOverwrite)

    #If CON_BLN_DebugModeActivated Then
        Debug.Print pSTR_SUBFCN_Name & " Performed"
    #End If

End Sub

Private Sub SUB_SeteIDVars(RNG_eIDTarget As Range, STR_1040eID As String, WS_1040K1Summ As Worksheet, _
                           RNG_1040eIDLookup As Range)
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To set validated eID variables for use in other procedures.

    pSTR_SUBFCN_Name = "SUB_SeteIDVars"

    STR_1040eID = RNG_eIDTarget.Text
    Set RNG_1040eIDLookup = WS_1040K1Summ.Range("I1:I878")

End Sub

Private Function fBLN_eIDExists(STR_1040eID As String)
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: Check to see if the eID is blank, if so return false.

    pSTR_SUBFCN_Name = "fBLN_eIDExists"

    If STR_1040eID = "" Then

        fBLN_eIDExists = False

    Else

        fBLN_eIDExists = True

    End If

End Function

Private Function fBLN_IseIDValid(BLN_eIDErrExists As Boolean, STR_1040eID As String, WS_1040K1Summ As Worksheet, _
                                 RNG_1040eIDLookup As Range, OBJ_ExpLog As Object) As Boolean
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: Check to see if the eID is in the eID list, if not, then return false.

    pSTR_SUBFCN_Name = "fBLN_IseIDValid"
    pSUB_SetDynamicStatusMsg STR_Msg:="Checking vailidity of Entity ID for " & STR_1040eID & "."

    If IsError(Application.Match(STR_1040eID, RNG_1040eIDLookup, 0)) Then

        fBLN_IseIDValid = False
        BLN_eIDErrExists = True
        OBJ_ExpLog.Write "ERROR DETECTED!" & " The Entity ID " & STR_1040eID & " is not valid." & vbNewLine & _
                         "Please confirm your input, the Entity ID in question has been marked red.'. " & vbNewLine & vbNewLine
        RNG_eIDTarget.Font.Color = RGB(255, 0, 0)

    Else

        fBLN_IseIDValid = True
        RNG_eIDTarget.Font.Color = RGB(255, 255, 255) ' Reset Color as needed

    End If

    #If CON_BLN_DebugModeActivated Then
        Debug.Print pSTR_SUBFCN_Name & ": " & STR_1040eID & " = " & fBLN_IseIDValid
    #End If

End Function

Private Function fSTR_K1OutPath(STR_1040eID As String, WS_eID As Worksheet, RNG_eIDLookup As Range, _
                                RNG_eIDPathLookup As Range) As String
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To lookup and return the K-1 Output File Path as String.

    pSTR_SUBFCN_Name = "fSTR_K1OutPath"
    Set RNG_eIDLookup = WS_eID.Range("B1:B100")
    Set RNG_eIDPathLookup = WS_eID.Range("C1:C100")

    With Application.WorksheetFunction

        Dim LNG_LookupRow As Long
        LNG_LookupRow = .Match(STR_1040eID, RNG_eIDLookup, 0)    ' Using helper variable because of issues with directly returning Lookup Row
        fSTR_K1OutPath = .Index(RNG_eIDPathLookup, LNG_LookupRow)

    End With

    #If CON_BLN_DebugModeActivated Then
        Debug.Print pSTR_SUBFCN_Name & " = " & fSTR_K1OutPath
    #End If

End Function

Private Function fBLN_K1OutPathExists(BLN_PathErrExists As Boolean, STR_1040eID As String, STR_K1OutPath As String, _
                                      OBJ_ExpLog As Object) As Boolean
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: Look up the file path from the ENTITY_ID worksheet based on Entity ID, if no file path, then skip to next.

    pSTR_SUBFCN_Name = "fBLN_K1OutPathExists"

    If STR_K1OutPath = "" Then

        fBLN_K1OutPathExists = False
        BLN_PathErrExists = True
        OBJ_ExpLog.Write "No File Path Found for Entity ID: " & STR_1040eID & "." & vbNewLine & _
                         "This is likely due to the fact that K-1 Output for this entity is not " & _
                         "yet ready for use, please confirm with preparer of this entity if you think this is in error." & _
                         vbNewLine & vbNewLine
    Else

        fBLN_K1OutPathExists = True

    End If

    #If CON_BLN_DebugModeActivated Then
        Debug.Print pSTR_SUBFCN_Name & ": " & STR_1040eID & " = " & fBLN_K1OutPathExists
    #End If


End Function

Private Function fBLN_K1OutPathIsValid(BLN_InvalidPathExists As Boolean, STR_1040eID As String, STR_K1OutPath As String, _
                                       OBJ_ExpLog As Object) As Boolean
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To determine if a K1 Output file path is valid or not.

    pSTR_SUBFCN_Name = "fBLN_K1OutPathIsValid"
    pSUB_SetDynamicStatusMsg STR_Msg:="Checking validity of K-1 Output File Path for " & STR_K1OutPath & "."
    
    If pfBLN_FilePathExists(STR_K1OutPath, BLN_DispMsg:=False) = False Then

        fBLN_K1OutPathIsValid = False
        BLN_InvalidPathExists = True
        OBJ_ExpLog.Write "ERROR DETECTED! The file path " & STR_K1OutPath & " for the Entity ID " & STR_1040eID & _
                         " is invalid, please confirm this input with the preparer." & vbNewLine & vbNewLine

    Else

        fBLN_K1OutPathIsValid = True

    End If

    #If CON_BLN_DebugModeActivated Then
        Debug.Print pSTR_SUBFCN_Name & ": " & STR_K1OutPath & " = " & fBLN_K1OutPathIsValid
    #End If

End Function

Private Sub SUB_SetpIDVars(RNG_eIDTarget As Range, STR_1040pID As String, WB_K1Out As Workbook, _
                           WS_K1Out As Worksheet, STR_K1OutPath As String, RNG_K1OutpIDLookup As Range, _
                           STR_1040pName As String)
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To set validated pID variables for use in other procedures.

    pSTR_SUBFCN_Name = "SUB_SetpIDVars"
    pSUB_SetDynamicStatusMsg STR_Msg:="Opening the K-1 Output located at " & STR_K1OutPath & "."

    STR_1040pID = RNG_eIDTarget.Offset(0, 1).Text
    Set WB_K1Out = Workbooks.Open(Filename:=STR_K1OutPath, ReadOnly:=True)
    Set WS_K1Out = WB_K1Out.Sheets(fSTR_SetK1OutEstOrAct(WB_K1Out))
    Set RNG_K1OutpIDLookup = WS_K1Out.Range("J1:J878")
    STR_1040pName = RNG_eIDTarget.Offset(0, -7).Text

    #If CON_BLN_DebugModeActivated Then
        Debug.Print pSTR_SUBFCN_Name & " Performed"
    #End If

End Sub

Private Function fBLN_pIDIsValid(BLN_pIDErrExists As Boolean, STR_1040pID As String, RNG_K1OutpIDLookup As Range, _
                                 STR_1040pName As String, OBJ_ExpLog As Object) As Boolean
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: Check to see if partner_id exists, if not write to error log and skip.

    pSTR_SUBFCN_Name = "fBLN_pIDIsValid"
    pSUB_SetDynamicStatusMsg STR_Msg:="Checking Validity of Partner ID for " & STR_1040pID & "."
    
    If IsError(Application.Match(STR_1040pID, RNG_K1OutpIDLookup, 0)) Then

        fBLN_pIDIsValid = False
        BLN_pIDErrExists = True
        OBJ_ExpLog.Write "ERROR DETECTED! " & STR_1040pName & " was NOT imported. The Partner ID of " & STR_1040pID & _
                         " does not exist in the target K-1 Output Sheet. Please confirm your input." & vbNewLine & vbNewLine
    Else

        fBLN_pIDIsValid = True

    End If
    #If CON_BLN_DebugModeActivated Then
        Debug.Print pSTR_SUBFCN_Name & ": " & STR_1040pID & " = " & fBLN_pIDIsValid
    #End If

End Function

Private Sub SUB_SetDateVars(STR_1040pID As String, RNG_1040pIDLookup As Range, RNG_K1OutpIDLookup As Range, _
                            LNG_1040ImportRow As Long, WS_1040K1Summ As Worksheet, RNG_1040Date As Range, _
                            LNG_K1OutImportRow As Long, WS_K1Out As Worksheet, RNG_K1OutDate As Range)
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To set validated date variables for use in other procedures.

    pSTR_SUBFCN_Name = "SUB_SetDateVars"

    Set RNG_1040pIDLookup = WS_1040K1Summ.Range("J1:J878")

    LNG_1040ImportRow = Application.WorksheetFunction.Match(STR_1040pID, RNG_1040pIDLookup, 0)
    Set RNG_1040Date = WS_1040K1Summ.Cells(LNG_1040ImportRow, fSTR_K1Col("DateCol"))

    LNG_K1OutImportRow = Application.WorksheetFunction.Match(STR_1040pID, RNG_K1OutpIDLookup, 0)
    Set RNG_K1OutDate = WS_K1Out.Cells(LNG_K1OutImportRow, fSTR_K1Col("DateCol"))

    #If CON_BLN_DebugModeActivated Then
        Debug.Print pSTR_SUBFCN_Name & " Performed"
    #End If


End Sub

Private Function fBLN_DateIsNotCurrent(STR_1040pID As String, RNG_1040Date As Range, RNG_K1OutDate As Range, _
                                       STR_1040pName As String, OBJ_ExpLog As Object) As Boolean
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: Check to see if the date stamp is the same as the source K-1 Output sheet, if it isn't or is blank
    '         then import, if it is, then don't import.

    pSTR_SUBFCN_Name = "fBLN_DateIsNotCurrent"
    pSUB_SetDynamicStatusMsg STR_Msg:="Checking if date stamp is current or not for " & STR_1040pID & "."
    
    If IsDate(RNG_1040Date.Value) Then

        If DateValue(RNG_1040Date.Value) = DateValue(RNG_K1OutDate.Value) Then

            fBLN_DateIsNotCurrent = False
            OBJ_ExpLog.Write STR_1040pName & " was NOT imported with Partner ID of " & STR_1040pID & _
                             " because the timestamp of this partner matched that of the K-1 Output for this entity." & _
                             vbNewLine & vbNewLine

        End If

    Else

        fBLN_DateIsNotCurrent = True

    End If

    #If CON_BLN_DebugModeActivated Then
        Debug.Print pSTR_SUBFCN_Name & ": " & STR_1040pID & " = " & fBLN_DateIsNotCurrent
    #End If

End Function

Private Sub SUB_SetImportVars(RNG_1040Import As Range, WS_1040K1Summ As Worksheet, LNG_1040ImportRow As Long, _
                              RNG_K1OutExport As Range, WS_K1Out As Worksheet, LNG_K1OutImportRow As Long, _
                              RNG_1040EstorAct As Range, RNG_K1OutEstorAct As Range)
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To set validated Import variables for use in other procedures.

    Set RNG_1040Import = WS_1040K1Summ.Range(WS_1040K1Summ.Cells(LNG_1040ImportRow, fSTR_K1Col("FirstImportCol")) _
                                             , WS_1040K1Summ.Cells(LNG_1040ImportRow, fSTR_K1Col("LastImportCol")))

    Set RNG_K1OutExport = WS_K1Out.Range(WS_K1Out.Cells(LNG_K1OutImportRow, fSTR_K1Col("FirstImportCol")) _
                                         , WS_K1Out.Cells(LNG_K1OutImportRow, fSTR_K1Col("LastImportCol")))

    Set RNG_1040EstorAct = WS_1040K1Summ.Cells(LNG_1040ImportRow, fSTR_K1Col("ESTorACTCol"))
    Set RNG_K1OutEstorAct = WS_K1Out.Cells(LNG_K1OutImportRow, fSTR_K1Col("ESTorACTCol"))

    #If CON_BLN_DebugModeActivated Then
        Debug.Print pSTR_SUBFCN_Name & " Performed"
    #End If

End Sub

Private Sub SUB_PerformImport(RNG_1040Import As Range, RNG_K1OutExport As Range, RNG_1040Date As Range, _
                              RNG_K1OutDate As Range, RNG_1040EstorAct As Range, RNG_K1OutEstorAct As Range, _
                              STR_1040pName As String)
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To actually perform the import of the values.
    
    pSTR_SUBFCN_Name = "SUB_PerformImport"
    pSUB_SetDynamicStatusMsg STR_Msg:="Importing K-1 Data for " & STR_1040pName & "."
    
    RNG_1040Import.Value = RNG_K1OutExport.Value    'Where the magic happens.
    RNG_1040Date.Value = RNG_K1OutDate.Value
    RNG_1040EstorAct.Value = RNG_K1OutEstorAct.Value
    pINT_Counter = pINT_Counter + 1

    #If CON_BLN_DebugModeActivated Then
        Debug.Print pSTR_SUBFCN_Name & " Performed"
    #End If

End Sub

Private Sub SUB_SetTaxIncVars(WS_1040K1Summ As Worksheet, RNG_1040TaxInc As Range, LNG_1040ImportRow As Long, _
                              WS_K1Out As Worksheet, RNG_K1OutTaxInc As Range, LNG_K1OutImportRow As Long)
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To set validated TaxInc variables for use in other procedures.

    pSTR_SUBFCN_Name = "SUB_SetTaxIncVars"

    Set RNG_1040TaxInc = WS_1040K1Summ.Cells(LNG_1040ImportRow, fSTR_K1Col("TaxableIncomeCol"))
    Set RNG_K1OutTaxInc = WS_K1Out.Cells(LNG_K1OutImportRow, fSTR_K1Col("TaxableIncomeCol"))

    #If CON_BLN_DebugModeActivated Then
        Debug.Print pSTR_SUBFCN_Name & " Performed"
    #End If

End Sub

Private Sub SUB_CheckTaxInc(BLN_TaxIncErrExists As Boolean, WS_1040K1Summ As Worksheet, RNG_1040TaxInc As Range, _
                            WS_K1Out As Worksheet, RNG_K1OutTaxInc As Range, STR_1040pName As String, _
                            RNG_1040Import As Range, STR_1040pID As String, OBJ_ExpLog As Object)
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To confirm taxable income has been imported correctly

    pSTR_SUBFCN_Name = "SUB_CheckTaxInc"
    pSUB_SetDynamicStatusMsg STR_Msg:="Verifying taxable income for " & STR_1040pName & "."

    Application.Calculation = xlCalculationAutomatic    'Need to turn on calculation to check taxable income.

    If RNG_1040TaxInc.Value = RNG_K1OutTaxInc.Value Then

        OBJ_ExpLog.Write STR_1040pName & " was imported into range " & RNG_1040Import.Address & " with Partner ID of " & _
                         STR_1040pID & ". Taxable Income for this partner was " & Format(RNG_1040TaxInc.Value, "Standard") & _
                         "." & vbNewLine & vbNewLine

    Else

        BLN_TaxIncErrExists = True
        OBJ_ExpLog.Write "ERROR DETECTED!" & STR_1040pName & " was imported into range " & RNG_1040Import.Address & _
                         " with Partner ID of " & STR_1040pID & "." & vbNewLine & "Taxable Income, which was in error, for this partner was " & _
                         Format(RNG_1040TaxInc.Value, "Standard") & ". Please review your Taxable Income computation on your 1040 Workpaper Set" & _
                         vbNewLine & vbNewLine

    End If

    Application.Calculation = xlCalculationManual

    #If CON_BLN_DebugModeActivated Then
        Debug.Print pSTR_SUBFCN_Name & "Performed"
    #End If

End Sub

Private Function fSTR_SetK1OutEstOrAct(WB_K1Out As Workbook) As String
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To support the Auto-import function

    pSTR_SUBFCN_Name = "fSTR_SetK1OutEstOrAct(WB_K1Out)"
    WB_K1Out.Activate

    If pfBLN_SheetExists(STR_SheetName:="K-1 OUTPUT", BLN_DispMsg:=False) Then

        fSTR_SetK1OutEstOrAct = "K-1 OUTPUT"

    ElseIf pfBLN_SheetExists(STR_SheetName:="K-1 OUTPUT - EST", BLN_DispMsg:=False) Then

        fSTR_SetK1OutEstOrAct = "K-1 OUTPUT - EST"

    Else

        MsgBox "No K-1 Output Sheet Exists in the Export Workbook " & WB_K1Out.Name & "!", vbCritical, "Error!"
        fSTR_SetK1OutEstOrAct = "ERROR"

    End If

    #If CON_BLN_DebugModeActivated Then
        Debug.Print pSTR_SUBFCN_Name & " = " & fSTR_SetK1OutEstOrAct
    #End If

End Function

Private Sub SUB_CompletionMsg(BLN_pIDErrExists As Boolean, BLN_eIDErrExists As Boolean, BLN_TaxIncErrExists As Boolean, _
                              BLN_PathErrExists As Boolean, BLN_InvalidPathExists As Boolean, STR_ExpLogPath As String, _
                              SGL_TimeElapsed As Single)
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To display completion message and allow user to open export log

    MsgBox "All available K-1 values have been imported, please check the export log to see what was imported." & vbNewLine & vbNewLine & _
           pINT_Counter & " K-1 value(s) have been imported into this 1040 K-1 Summary." & vbNewLine & vbNewLine & "This process took " & _
           SGL_TimeElapsed & " seconds to complete." & vbNewLine & vbNewLine & fSTR_pIDErrMsg(BLN_pIDErrExists) & _
           vbNewLine & vbNewLine & fSTR_eIDErrMsg(BLN_eIDErrExists) & vbNewLine & vbNewLine & fSTR_TaxIncErrMsg(BLN_TaxIncErrExists) & vbNewLine & vbNewLine & _
           fSTR_PathErrMsg(BLN_PathErrExists) & vbNewLine & vbNewLine & fSTR_InvalidPathErrMsg(BLN_InvalidPathExists), vbOKOnly, "Automatic Import of K-1 Values"

    pSTR_SUBFCN_Name = "SUB_CompletionMsg"

    If MsgBox("Would you like to open the export log?", vbYesNo, "Open Export Log") = vbYes Then

        Dim DBL_Export_Log As Double
        DBL_Export_Log = Shell("C:\WINDOWS\notepad.exe " & STR_ExpLogPath, vbNormalFocus)

    End If

End Sub

Private Function fSTR_pIDErrMsg(BLN_pIDErrExists As Boolean) As String
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To support the Auto-import function

    pSTR_SUBFCN_Name = "fSTR_pIDErrMsg"

    If BLN_pIDErrExists Then

        fSTR_pIDErrMsg = "There was a Partner ID error(s) detected, please check the export log and confirm applicable partner numbers and re-import if needed"

    Else

        fSTR_pIDErrMsg = "There were no Partner ID error(s) detected."

    End If

End Function

Private Function fSTR_eIDErrMsg(BLN_eIDErrExists As Boolean) As String
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To support the Auto-import function

    pSTR_SUBFCN_Name = "fSTR_eIDErrMsg"

    If BLN_eIDErrExists Then

        fSTR_eIDErrMsg = "There was an Entity ID error(s) detected in this 1040 K-1 Summary, please re-run the 'Validate Entity IDs' Tool and confirm your input."

    Else

        fSTR_eIDErrMsg = "There were no Entity ID error(s) detected."

    End If

End Function

Private Function fSTR_TaxIncErrMsg(BLN_TaxIncErrExists As Boolean) As String
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To support the Auto-import function

    pSTR_SUBFCN_Name = "fSTR_TaxIncErrMsg"

    If BLN_TaxIncErrExists Then

        fSTR_TaxIncErrMsg = "There was an Taxable Income error(s) detected, please exit without saving, and check the export log, then confirm the calculation."

    Else

        fSTR_TaxIncErrMsg = "There were no Taxable Income error(s) detected."

    End If

End Function

Private Function fSTR_PathErrMsg(BLN_PathErrExists As Boolean) As String
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To support the Auto-import function

    pSTR_SUBFCN_Name = "fSTR_PathErrMsg"

    If BLN_PathErrExists Then

        fSTR_PathErrMsg = "There were Entity IDs in this workbook for which no corresponding K-1 Output was found, check the export log for detail." & _
                          " They are likely not yet prepared for use."

    Else

        fSTR_PathErrMsg = "All Entity IDs had associated K-1 Output files ready for use."

    End If

End Function

Private Function fSTR_InvalidPathErrMsg(BLN_InvalidPathExists As Boolean) As String
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To support the Auto-import function

    pSTR_SUBFCN_Name = "fSTR_InvalidPathErrMsg"

    If BLN_InvalidPathExists Then

        fSTR_InvalidPathErrMsg = "There exists an invalid file path in the Entity ID workbook, please confirm the file path with the preparer."

    Else

        fSTR_InvalidPathErrMsg = "All file paths used in this import were valid."

    End If

End Function













