Attribute VB_Name = "MOD__pVAR_pFCN_pSUB"
Option Explicit
'Purpose: To declare various public variables, functions, and subs for the Tax Excel Addin

'DEBUGGING NOTE: This project makes use of a conditional compilation constant which is declared in the general properties of
'                this project. The variable is "CON_CON_BLN_DebugModeActivated". You set the variables with 1 or 0 in the properties.

'Declare Public Variable for most recent version number for 1040 Workpaper
Public Const pSTR_CurrentVersion_1040 As String = "VERSION 1.10"

'Declare Public Variable for most recent version number for Entity K-1 Output Tab
'Should generally be the same as the 1040 number, but in case in future different numbers are desired
Public Const pSTR_CurrentVersion_Entity As String = "VERSION 1.7"

'Declare Public Variable for most recent version number for the Tax Excel Add-in itself.
Public Const pSTR_CurrentVersion_Addin As String = "VERSION 2.1.1"

'Need this type to call windows API to get millisecond accurate system time used in the pfSTR_ReturnSystemTime
Public Type SYSTEMTIME
    pINT_Year As Integer
    pINT_Month As Integer
    pINT_DayOfWeek As Integer
    pINT_Day As Integer
    pINT_Hour As Integer
    pINT_Minute As Integer
    pINT_Second As Integer
    pINT_Milliseconds As Integer
End Type

Public Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

'Global Error Handler
Public pSTR_SUBFCN_Name As String
Public pBLN_WorkpaperUpToDate As Boolean

'Declare iomode values for FSO textream values to be used elsewhere.
Public Const fsoForReading = 1
Public Const fsoForWriting = 2
Public Const fsoForAppending = 8
Public Const fsoForOverwrite = True

'Declare generic variables for use in various subs to avoid activesheet/activeworkbook
Public pWB_User As Workbook
Public pWS_User As Worksheet
Public pRNG_Target As Range
Public pRNG_TargetRange As Range
Public pPIC_Target As Picture
Public pSHP_Target As Shape
Public pVAR_Target As Variant
Public pINT_Target As Integer
Public pINT_Counter As Integer
Public pWB_Template As Workbook

Public Sub pSUB_OpenTemplate(STR_TemplateName As String)
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To open certain excel templates based on STR_TemplateName

    pSTR_SUBFCN_Name = "pSUB_OpenTemplate"
    On Error GoTo ErrorHandler

    Dim STR_FilePath As String
    Select Case STR_TemplateName

        Case "Template1040_CurrentVersion"
            STR_FilePath = "J:\TAX\Tax Excel Add-in\Template_1040\Template_1040_Current_Version.xlsx"
        Case "Template1040_1_2"
            STR_FilePath = "J:\TAX\Tax Excel Add-in\Template_1040\Template_1040_1_2.xlsx"
        Case "Template1040_1_3"
            STR_FilePath = "J:\TAX\Tax Excel Add-in\Template_1040\Template_1040_1_3.xlsx"
        Case "Template1040_1_4"
            STR_FilePath = "J:\TAX\Tax Excel Add-in\Template_1040\Template_1040_1_4.xlsx"
        Case "Template1040_1_5"
            STR_FilePath = "J:\TAX\Tax Excel Add-in\Template_1040\Template_1040_1_5.xlsx"
        Case "Template1040_1_6"
            STR_FilePath = "J:\TAX\Tax Excel Add-in\Template_1040\Template_1040_1_6.xlsx"
        Case "Template1040_1_7"
            STR_FilePath = "J:\TAX\Tax Excel Add-in\Template_1040\Template_1040_1_7.xlsx"
        Case "Template1040_1_8"
            STR_FilePath = "J:\TAX\Tax Excel Add-in\Template_1040\Template_1040_1_8.xlsx"

        Case Else
            #If CON_BLN_DebugModeActivated Then
                Debug.Print "STR_TemplateName not found, add to case list"
            #End If

    End Select

    Set pWB_Template = Workbooks.Open(Filename:=STR_FilePath, ReadOnly:=True)

    On Error GoTo 0
    Exit Sub
ErrorHandler:
    pSUB_GlobalErrorHandler

End Sub

Public Sub pSUB_PerformanceOptions(BLN_TurnOn As Boolean)
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To house standard performance increasing and interference reducing options for subs in the Add-in
    'Note: BLN_TurnOn, True at start of sub, False at end of sub.
    'Note: Not giving this function a name for error handling purposes to avoid having excessive renames of that variable.
    'Note: This also triggers a Modeless Userform during code execution to display status messages to user.

    Select Case BLN_TurnOn

        Case True
            DoEvents    ' To force refresh

            Dim STR_StaticStatusMsg As String
            STR_StaticStatusMsg = "Processing Current Request... Please Wait."

            With FRM_KTCTaxAddinStatus
                .StartUpPosition = 0
                .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)    ' To ensure center placement
                .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)    ' To ensure center placement
                .Show
                .LBL_StaticStatusMsg = STR_StaticStatusMsg
                .LBL_DynamicStatusMsg = ""    'Clear any old value
                .Repaint
            End With

            With Application
                .StatusBar = STR_StaticStatusMsg
                .Cursor = xlWait
                .EnableEvents = False
                .ScreenUpdating = False
                .Calculation = xlCalculationManual
                .DisplayAlerts = False
            End With

        Case False

            FRM_KTCTaxAddinStatus.Hide

            With Application
                .ScreenUpdating = True
                .Calculation = xlCalculationAutomatic
                .EnableEvents = True
                .StatusBar = False
                .Cursor = xlDefault
                .DisplayAlerts = True
            End With

    End Select

End Sub

Public Sub pSUB_SetDynamicStatusMsg(STR_Msg As String)
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To update the dynamic status message for use in the Modeless userform that displays during code execution

    With FRM_KTCTaxAddinStatus
        .LBL_DynamicStatusMsg = STR_Msg
        .Repaint
    End With

End Sub

Public Function pfBLN_SheetExists(STR_SheetName As String, BLN_DispMsg As Boolean) As Boolean
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To determine if the specified worksheet exists in the workbook. Will throw message if sheet doesn't exist.

    pSTR_SUBFCN_Name = "pfBLN_SheetExists"
    On Error GoTo ErrorHandler
    Dim WS_Target As Worksheet
    pfBLN_SheetExists = False

    For Each WS_Target In ActiveWorkbook.Worksheets

        If STR_SheetName = WS_Target.Name Then

            pfBLN_SheetExists = True

            Exit For
        End If
    Next

    If pfBLN_SheetExists = False And BLN_DispMsg = True Then
        MsgBox "The required sheet " & STR_SheetName & " does not exist in this workbook.", vbCritical, "Error!"
    End If

    #If CON_BLN_DebugModeActivated Then
        Debug.Print "Required sheet " & STR_SheetName & " does not exist."
    #End If

    On Error GoTo 0
    Exit Function
ErrorHandler:
    pSUB_GlobalErrorHandler

End Function

Public Function pfBLN_ActiveSheetExists(BLN_DispMsg As Boolean) As Boolean
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To determine if any sheet is active, return true.
    'Note: Need to pass string as 'DisplayMsgYes' or 'DisplayMsgNo' for error message.

    On Error GoTo ErrorHandler

    Dim STR_TempNameHolder As String
    STR_TempNameHolder = ActiveSheet.Name
    pfBLN_ActiveSheetExists = True

    #If CON_BLN_DebugModeActivated Then
        Debug.Print "pfBLN_ActiveSheetExists = " & pfBLN_ActiveSheetExists
    #End If

    On Error GoTo 0
    Exit Function
ErrorHandler:
    'Not using global error handler here becuase the error handler is part of the function, relying on error to actually do the testing.

    pfBLN_ActiveSheetExists = False
    If BLN_DispMsg = True Then MsgBox "No open worksheets found, this tool will not run.", vbCritical, "Error!"

    #If CON_BLN_DebugModeActivated Then
        Debug.Print "pfBLN_ActiveSheetExists = " & pfBLN_ActiveSheetExists
    #End If

End Function

Public Function pfSTR_ReturnSystemTime() As String
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To return millisecond accurate system time using the windows API

    On Error GoTo ErrorHandler

    Dim TYP_System As SYSTEMTIME
    GetSystemTime TYP_System
    pfSTR_ReturnSystemTime = Month(Now) & ":" & Day(Now) & ":" & Year(Now) & ":" & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now) & ":" & TYP_System.pINT_Milliseconds

    On Error GoTo 0
    Exit Function
ErrorHandler:
    pSUB_GlobalErrorHandler

End Function

Public Function pfARR_LoadSheetNames() As Variant
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To load all sheet names to an array

    ReDim VAR_SheetNames(1 To ActiveWorkbook.Sheets.Count)

    Dim INT_Target As Integer
    For INT_Target = 1 To Sheets.Count
        VAR_SheetNames(INT_Target) = Sheets(INT_Target).Name
    Next

    pfARR_LoadSheetNames = VAR_SheetNames

End Function

Public Sub pSUB_TimeElapsed(SGL_StartTime As Single, SGL_TimeElapsed As Single, BLN_StartTimer As Boolean)
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To calculate the seconds elapsed
    'Note: Need to call this function twice, once at begining of procedure, with
    '      BLN_StartTimer as True and then at the end as False

    Select Case BLN_StartTimer

        Case True
            SGL_StartTime = Timer
            SGL_TimeElapsed = 0

        Case False
            SGL_TimeElapsed = Round(Timer - SGL_StartTime, 0)

    End Select

End Sub

Public Function pfBLN_FilePathExists(STR_FilePath As String, BLN_DispMsg As Boolean) As Boolean
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To determine if a file path is actually in existence

    Dim OBJ_FSO As Object
    Set OBJ_FSO = CreateObject("Scripting.FileSystemObject")

    If OBJ_FSO.FileExists(STR_FilePath) Then

        pfBLN_FilePathExists = True

    Else

        pfBLN_FilePathExists = False
        If BLN_DispMsg = True Then MsgBox "The required file " & STR_FilePath & " does not exist. Please confirm input or contact developer.", vbCritical, "File Path Not Found!"

    End If

    #If CON_BLN_DebugModeActivated Then
        Debug.Print "pfBLN_FilePathExists: " & STR_FilePath & " = " & pfBLN_FilePathExists
    #End If

End Function

Public Function pfBLN_IsSelectionRange(BLN_DispMsg As Boolean) As Boolean
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To determine if the selected object is a range.

    If pfBLN_ActiveSheetExists(BLN_DispMsg:=False) = False Then Exit Function

    pSTR_SUBFCN_Name = "pfBLN_IsSelectionRange"
    On Error GoTo ErrorHandler

    Select Case TypeName(Selection)

        Case "Range"
            pfBLN_IsSelectionRange = True

        Case Else
            pfBLN_IsSelectionRange = False
            If BLN_DispMsg = True Then

                MsgBox "You have selected something other than a cell range, this tool will not function. " & vbNewLine & vbNewLine & _
                       "Please select a cell or range of cells and try again", vbCritical, "Error! Range not selected!"

            End If

    End Select

    #If CON_BLN_DebugModeActivated Then
        Debug.Print "pfBLN_IsSelectionRange = " & pfBLN_IsSelectionRange
    #End If

    On Error GoTo 0
    Exit Function
ErrorHandler:
    pSUB_GlobalErrorHandler

End Function

Public Sub pSUB_DeleteAllShapes()
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To delete all shapes in a workbook without regard to type or name.

    Set pWB_User = ActiveWorkbook

    For Each pWS_User In pWB_User.Sheets

        If pWS_User.Shapes.Count > 0 Then

            For Each pSHP_Target In pWS_User.Shapes

                pSHP_Target.Delete

            Next
        End If
    Next

    #If CON_BLN_DebugModeActivated Then
        Debug.Print pSTR_SUBFCN_Name & " Performed"
    #End If

End Sub

Public Function pfARR_RangetoStringArray(RNG_RangetoConvert As Range) As String()

    Dim ARR_TempVariantArray As Variant
    Dim ARR_TempStringArray() As String

    ARR_TempVariantArray = RNG_RangetoConvert.Value
    ReDim ARR_TempStringArray(1 To UBound(ARR_TempVariantArray))

    For pINT_Target = 1 To UBound(ARR_TempVariantArray)
        
        ARR_TempStringArray(pINT_Target) = ARR_TempVariantArray(pINT_Target, 1)
    
    Next

    pfARR_RangetoStringArray = ARR_TempStringArray()

End Function







