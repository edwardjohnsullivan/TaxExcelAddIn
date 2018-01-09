Attribute VB_Name = "MOD__GlobalErrHandler"
Option Explicit
Public Sub pSUB_GlobalErrorHandler()
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com   
    'Purpose: To provide global error catch for subs in the Tax Excel Add-in
    '         Will write error to a error log on the network, provide a msg box, and a console message

    Dim STR_ErrMsg As String
    STR_ErrMsg = "An unexpected error has occurred at " & Now & vbNewLine & _
                 "The user who experienced the error was " & Application.UserName & vbNewLine & _
                 "The error was caused by " & pSTR_SUBFCN_Name & vbNewLine & _
                 "The error number is " & Err.Number & vbNewLine & _
                 "The error description is " & Err.Description & vbNewLine & vbNewLine

    Dim OBJ_FSO As Object
    Set OBJ_FSO = CreateObject("Scripting.FileSystemObject")

    Dim STR_GlobalErrLog_Path As String
    STR_GlobalErrLog_Path = "J:\TAX\Tax Excel Add-in\Logs\GlobalErrorLog.txt"
    Dim OBJ_GlobalErrLog_File As Object
    Set OBJ_GlobalErrLog_File = OBJ_FSO.OpenTextFile(STR_GlobalErrLog_Path, fsoForAppending)

    With OBJ_GlobalErrLog_File
        .Write STR_ErrMsg
        .Close
    End With

    Debug.Print STR_ErrMsg

    MsgBox "An unexpected error has occurred, please retry." & vbNewLine & _
           "If error persists, contact developer.", vbCritical, "Unexpected Error!"

    If pfBLN_ActiveSheetExists(BLN_DispMsg:=False) = True Then pSUB_PerformanceOptions BLN_TurnOn:=False
    ' It is likely that pSUB_PerformanceOptionsBLN_TurnOn:=True has been set if there was an unexpected error.

End Sub

