Attribute VB_Name = "MOD_WpVersCheck"
Public Function pfBLN_IsWpUpToDate(STR_WpName As String, BLN_DispUpdateMsg As Boolean, BLN_DispErrMsg As Boolean) As Boolean
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: Automatic Version Check of Tax 1040 Workpapers
    'Note: BLN_DispMsg for this sub is only the affirmation of an up to date workpaper,
    '      an out of date workpaper will always display a message.

    If pfBLN_ActiveSheetExists(BLN_DispMsg:=False) = False Then Exit Function

    pSTR_SUBFCN_Name = "pfBLN_IsWpUpToDate"
    On Error GoTo ErrorHandler

    Set pWB_User = ActiveWorkbook
    If fBLN_IsWpCheckTarget = False Then Exit Function

    Dim RNG_VerNumber As Range
    Dim STR_Version As String

    Select Case STR_WpName

        Case "1040"
            If pfBLN_SheetExists(STR_SheetName:="GUIDE", BLN_DispMsg:=False) = False Then Exit Function
            If fBLN_IsWpCheckTarget = False Then Exit Function
            Set RNG_VerNumber = pWB_User.Sheets("GUIDE").Range("A2")
            STR_Version = pSTR_CurrentVersion_1040

        Case "K1Out"
            If pfBLN_SheetExists(STR_SheetName:="K-1 OUTPUT", BLN_DispMsg:=False) = False Then Exit Function
            If fBLN_IsWpCheckTarget = False Then Exit Function
            Set RNG_VerNumber = pWB_User.Sheets("K-1 OUTPUT").Range("A4")
            STR_Version = pSTR_CurrentVersion_Entity

    End Select

    Select Case RNG_VerNumber.Value

        Case Is = STR_Version
            pfBLN_IsWpUpToDate = True
            If BLN_DispUpdateMsg = True Then MsgBox "This " & STR_WpName & " is up to date, no action required.", vbOKOnly, "Version Check of " & STR_WpName

        Case Else
            pfBLN_IsWpUpToDate = False
            If BLN_DispErrMsg = True Then MsgBox "This " & STR_WpName & " is out of date, please update to " & STR_Version, vbCritical, "Automatic Version Check of " & STR_WpName

    End Select

    #If CON_BLN_DebugModeActivated Then
        Debug.Print STR_WpName & ", pBLN_WorkpaperUpToDate = " & pBLN_WorkpaperUpToDate
    #End If

    On Error GoTo 0
    Exit Function
ErrorHandler:
    pSUB_GlobalErrorHandler

End Function

Private Function fBLN_IsWpCheckTarget() As Boolean
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To determine if the workpaper is one you actually want to auto-check
    'Note: We only auto-check workpapers that have the current tax year, or the next tax year in the name.
    '      This is to avoid templates and old workpapers that we would not want to update.

    fBLN_IsWpCheckTarget = False

    If fBLN_IsWpTemplate = False And fBLN_IsWpApplicableYear = True Then fBLN_IsWpCheckTarget = True

    #If CON_BLN_DebugModeActivated Then
        Debug.Print "fBLN_IsWpCheckTarget = " & fBLN_IsWpCheckTarget
    #End If

End Function

Private Function fBLN_IsWpTemplate() As Boolean
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To determine if the workpaper is in the Templates Folder

    fBLN_IsWpTemplate = False

    If pWB_User.Path = "J:\TAX\Workpaper Templates" Then fBLN_IsWpTemplate = True

End Function

Private Function fBLN_IsWpApplicableYear() As Boolean
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To determine if the workpaper has an applicable tax year (Current year or one back) for updating.

    fBLN_IsWpApplicableYear = False

    Dim INT_ThisTaxYear As Integer
    Dim INT_NextTaxYear As Integer
    INT_ThisTaxYear = 16    ' Two digit year to follow Tax Naming Convention
    INT_NextTaxYear = 17    ' Two digit year to follow Tax Naming Convention

    Dim ARR_ValidWorkpaperNames(1) As String

    ARR_ValidWorkpaperNames(0) = "*" & INT_ThisTaxYear & "*"
    ARR_ValidWorkpaperNames(1) = "*" & INT_NextTaxYear & "*"

    For pINT_Target = LBound(ARR_ValidWorkpaperNames) To UBound(ARR_ValidWorkpaperNames)

        If pWB_User.Name Like ARR_ValidWorkpaperNames(pINT_Target) Then

            fBLN_IsWpApplicableYear = True
            Exit Function

        End If
    Next

End Function
