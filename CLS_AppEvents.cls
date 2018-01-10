VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLS_AppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' Author: Edward Sullivan edwardjohnsullivan@gmail.com
' Purpose: To store subs that rely on Workbook Open event.
' Description: This module plus code that is in "ThisWorkbook" will do the following:
'                   Set Custom Keyboard Shortcuts
'                   Set Custom Context Menu Items
'                   Automatically Check the Addin Version
'                   Automatically Check the 1040 Workpaper Set Version
'                   Automatically Check the K-1 Output Version

'Set event monitoring for excel application
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1


Public Sub Class_Initialize()
    'Intialize the Application Event Watching.
    'This is needed for version check on workpaper open to function.

    Set App = Application

End Sub

Public Sub App_WorkbookOpen(ByVal WB_Review_Target As Workbook)
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To execute various subs upon any opening of a workbook.

    SUB_SetKeyboardShorcuts
    SUB_SetContextMenu
    SUB_AutoCheckAddinVersion
    pfBLN_IsWpUpToDate STR_WpName:="1040", BLN_DispUpdateMsg:=False, BLN_DispErrMsg:=False
    pfBLN_IsWpUpToDate STR_WpName:="K1Out", BLN_DispUpdateMsg:=False, BLN_DispErrMsg:=False

End Sub

Private Sub SUB_SetKeyboardShorcuts()
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: Establish Keyboard Shortcuts

    Application.OnKey "^+U", "pSUB_UnHide_Select"
    Application.OnKey "^+H", "pSUB_Hide_Select"
    Application.OnKey "^+T", "pSUB_InsertCheckmarksRange"
    Application.OnKey "^+D", "pSUB_DeleteCheckmarksRange"

End Sub

Private Sub SUB_SetContextMenu()
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: Add custom cell context menu items

    Dim CBR_ContextMenu As CommandBar
    Dim CBR_Control As CommandBarControl

    Set CBR_ContextMenu = Application.CommandBars("Cell")
    SUB_DeleteContextMenu CBR_ContextMenu, CBR_Control

    With CBR_ContextMenu.Controls.Add(Type:=msoControlButton, before:=5)
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "pSUB_InsertCheckmarksRange"
        .FaceId = 16470    'Checkmark Icon
        .Caption = "Insert Check mark(s)"
        .Tag = "KTC_Control_Tag"    ' Don't change, needs to be constant becuase tag will exist in operating environment
    End With

    With CBR_ContextMenu.Controls.Add(Type:=msoControlButton, before:=6)
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "pSUB_DeleteCheckmarksRange"
        .FaceId = 16471    'X Icon
        .Caption = "Delete Check mark(s)"
        .Tag = "KTC_Control_Tag"    ' Don't change, needs to be constant becuase tag will exist in operating environment
    End With

    CBR_ContextMenu.Controls(5).BeginGroup = True

End Sub

Private Sub SUB_DeleteContextMenu(CBR_ContextMenu As CommandBar, CBR_Control As CommandBarControl)
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: Delete custom cell context menu items

    For Each CBR_Control In CBR_ContextMenu.Controls
        If CBR_Control.Tag = "KTC_Control_Tag" Then CBR_Control.Delete
    Next

End Sub

Private Sub SUB_AutoCheckAddinVersion()
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: Automatic Version Check of Tax Excel Addin

    pSTR_SUBFCN_Name = "SUB_AutoCheckAddinVersion"
    On Error GoTo ErrorHandler

    'See if txt file string version number matches version number of currently loaded addin do nothing if match, send message if not match.

    If Not StrComp(pSTR_CurrentVersion_Addin, fSTR_ExtCurVersAddin, vbTextCompare) = 0 Then

        MsgBox "This Tax Excel Add-in is out of date." & vbNewLine & vbNewLine & _
               "The Current Version is " & fSTR_ExtCurVersAddin & "." & vbNewLine & vbNewLine & _
               fSTR_ExtUpdateMsg & vbNewLine & vbNewLine & _
               "Please exit all open Excel workbooks to allow the Excel to reload the Current Version of the Add-in at your earliest convenience.", _
               vbCritical, "Tax Excel Add-in is out of date!"

        Debug.Print "This Tax Excel Add-in is not the most current version."
        On Error GoTo 0
        Exit Sub

    End If

    #If CON_BLN_DebugModeActivated Then
        Debug.Print "This Tax Excel Add-in is the most current version."
    #End If

    On Error GoTo 0
    Exit Sub
ErrorHandler:
    pSUB_GlobalErrorHandler

End Sub

Private Function fSTR_ExtCurVersAddin() As String
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: 'Load most current external version number of add-in from network stored text file

    Dim OBJ_FSO As Object
    Set OBJ_FSO = CreateObject("Scripting.FileSystemObject")

    Dim STR_ExtCurrVerAddin_Path As String
    STR_ExtCurrVerAddin_Path = "J:\TAX\Tax Excel Add-in\External_Variables\STR_Ext_Current_Version_Addin.txt"
    Dim OBJ_ExtCurrVerAddin_File As Object
    Set OBJ_ExtCurrVerAddin_File = OBJ_FSO.OpenTextFile(STR_ExtCurrVerAddin_Path, fsoForReading)
    fSTR_ExtCurVersAddin = OBJ_ExtCurrVerAddin_File.ReadAll
    OBJ_ExtCurrVerAddin_File.Close

End Function

Private Function fSTR_ExtUpdateMsg() As String
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: Load current version update message into a string for use in warning message

    Dim OBJ_FSO As Object
    Set OBJ_FSO = CreateObject("Scripting.FileSystemObject")

    Dim STR_ExtUpdateMsg_Path As String
    STR_ExtUpdateMsg_Path = "J:\TAX\Tax Excel Add-in\External_Variables\STR_Ext_Update_Message.txt"
    Dim OBJ_ExtUpdateMsg_File As Object
    Set OBJ_ExtUpdateMsg_File = OBJ_FSO.OpenTextFile(STR_ExtUpdateMsg_Path, fsoForReading)
    fSTR_ExtUpdateMsg = OBJ_ExtUpdateMsg_File.ReadAll
    OBJ_ExtUpdateMsg_File.Close

End Function

