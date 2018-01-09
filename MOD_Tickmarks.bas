Attribute VB_Name = "MOD_Tickmarks"
Option Explicit
'Purpose: To house various subs supporting the "INSERTING CUSTOM TICKMARKS" feature of the Tax Excel Add-in

Public Sub pSUB_InsertCheckmarksRange(Optional Control As IRibbonControl)
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: Will insert a properly colored checkmark into the current selection.
    'Note: Needs to be public because keyboard shortcut calls.

    If pfBLN_ActiveSheetExists(BLN_DispMsg:=True) = False Then Exit Sub
    If pfBLN_IsSelectionRange(BLN_DispMsg:=True) = False Then Exit Sub

    pSTR_SUBFCN_Name = "pSUB_InsertCheckmarksRange"
    On Error GoTo ErrorHandler

    pSUB_PerformanceOptions BLN_TurnOn:=True

    If fBLN_DisplayObjectsAllowed = True Then SUB_InsertCheckmarks

    pSUB_PerformanceOptions BLN_TurnOn:=False

    On Error GoTo 0
    Exit Sub
ErrorHandler:
    pSUB_GlobalErrorHandler

End Sub

Private Sub SUB_InsertCheckmarks()
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: Will insert a checkmark.

    Dim STR_CheckmarkPath As String
    Dim STR_CheckmarkName As String
    Dim LNG_PicLeftOffset As Long
    SUB_SetCheckVars STR_CheckmarkPath, STR_CheckmarkName, LNG_PicLeftOffset

    Set pRNG_TargetRange = Selection
    Set pWS_User = ActiveSheet

    For Each pRNG_Target In pRNG_TargetRange
        With pWS_User.Pictures.Insert(Filename:=STR_CheckmarkPath)
            With .ShapeRange
                .LockAspectRatio = msoTrue
                .Width = 9    'Small but not too small.
            End With
            .Name = STR_CheckmarkName & pfSTR_ReturnSystemTime    'Adding time stamp here so it changes at every insert, need millisecond accuracy for unique values
            .Left = pRNG_Target.Left + LNG_PicLeftOffset    ' Variable to be able to see different user levels check marks.
            .Top = pRNG_Target.Top + 1    ' Always 1 to get off the top border to avoid stretching from hidden rows.
            .Placement = xlMoveAndSize
            .PrintObject = True
        End With
    Next

End Sub

Private Function fBLN_DisplayObjectsAllowed() As Boolean
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To determine if the workbook level setting for display of objects is set as needed.

    pSTR_SUBFCN_Name = "fBLN_DisplayObjectsAllowed"
    Set pWB_User = ActiveWorkbook

    Select Case pWB_User.DisplayDrawingObjects

        Case xlDisplayShapes
            fBLN_DisplayObjectsAllowed = True

        Case Else
            If MsgBox("This workbook has been set to hide all objects, including checkmarks." & vbNewLine & vbNewLine & _
                      "Checkmarks can not be inserted without allowing the display of all objects." & vbNewLine & vbNewLine & _
                      "Would you like to automatically change this setting to allow the display of all objects?", vbYesNo + vbCritical, _
                      "Error! Object Display Disabled.") = vbYes Then

                pWB_User.DisplayDrawingObjects = xlDisplayShapes
                fBLN_DisplayObjectsAllowed = True

            Else

                fBLN_DisplayObjectsAllowed = False
                Exit Function

            End If
    End Select

End Function

Private Sub SUB_SetCheckVars(STR_CheckmarkPath As String, STR_CheckmarkName As String, LNG_PicLeftOffset As Long)
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: Sets path and name (less time stamp) of checkmark based on user level

    pSTR_SUBFCN_Name = "SUB_SetCheckVars"

    Dim User As New CLS_User

    Select Case User.STR_OrgLevel

        Case "Preparer"
            STR_CheckmarkPath = "J:\TAX\Tax Excel Add-in\Icons\Tickmarks\Red_Checkmark.PNG"
            STR_CheckmarkName = User.STR_Initial & "_RedChk_"
            LNG_PicLeftOffset = 1    'Needs to be 1 point to not touch cell border to avoid stretching on unhide
        Case "Reviewer"
            STR_CheckmarkPath = "J:\TAX\Tax Excel Add-in\Icons\Tickmarks\Green_Checkmark.PNG"
            STR_CheckmarkName = User.STR_Initial & "_GreenChk_"
            LNG_PicLeftOffset = 5    'Offset next level by 4 points
        Case "Director"
            STR_CheckmarkPath = "J:\TAX\Tax Excel Add-in\Icons\Tickmarks\Blue_Checkmark.PNG"
            STR_CheckmarkName = User.STR_Initial & "_BlueChk_"
            LNG_PicLeftOffset = 9    'Offset next level by 4 points
        Case "Trust_Dept"
            STR_CheckmarkPath = "J:\TAX\Tax Excel Add-in\Icons\Tickmarks\Purple_Checkmark.PNG"
            STR_CheckmarkName = User.STR_Initial & "_PurpleChk_"
            LNG_PicLeftOffset = 13    'Offset next level by 4 points
        Case "Error"
            Exit Sub

    End Select

End Sub

Public Sub pSUB_DeleteChecksOption(STR_DeleteWhere As String)
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: Will delete all checkmarks based on STR_DeleteWhere value
    'Note: Needs to be public because keyboard shortcut calls.

    If pfBLN_ActiveSheetExists(BLN_DispMsg:=True) = False Then Exit Sub

    pSTR_SUBFCN_Name = "pSUB_DeleteChecksOption"
    On Error GoTo ErrorHandler

    pSUB_PerformanceOptions BLN_TurnOn:=True
    Set pWS_User = ActiveSheet

    Select Case STR_DeleteWhere

        Case "Range"
            Select Case TypeName(Selection)

                Case "Picture"    'Sometimes user selects a checkmark itself, and clicks delete, handled here.
                    Set pPIC_Target = Selection
                    SUB_DeleteChecks STR_DeleteWhere:="Range"

                Case "Range"
                    Set pRNG_Target = Selection
                    For Each pPIC_Target In pWS_User.Pictures
                        If fBLN_IsCheckInRange = True Then SUB_DeleteChecks STR_DeleteWhere:="Range"
                    Next

            End Select

        Case "Sheet"
            SUB_DeleteChecks STR_DeleteWhere:="Sheet"

        Case "Workbook"
            SUB_DeleteChecks STR_DeleteWhere:="Workbook"

    End Select

    pSUB_PerformanceOptions BLN_TurnOn:=False

    On Error GoTo 0
    Exit Sub
ErrorHandler:
    pSUB_GlobalErrorHandler

End Sub

Private Sub SUB_DeleteChecks(STR_DeleteWhere As String)
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To actually peform the deletion of a checkmark based on STR_DeleteWhere value

    pSTR_SUBFCN_Name = "SUB_DeleteChecks"

    Select Case STR_DeleteWhere

        Case "Range"
            If fBLN_IsPicTickmark Then pPIC_Target.Delete

        Case "Sheet"
            For Each pPIC_Target In pWS_User.Pictures
                SUB_DeleteChecks STR_DeleteWhere:="Range"
            Next

        Case "Workbook"
            For Each pWS_User In pWB_User.Worksheets
                SUB_DeleteChecks STR_DeleteWhere:="Sheet"
            Next

    End Select

End Sub

Private Function fBLN_IsCheckInRange() As Boolean
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To detemine if a picture is in the selected range.

    pSTR_SUBFCN_Name = "fBLN_IsCheckInRange"
    fBLN_IsCheckInRange = False

    If Not Intersect(pRNG_Target, Range(pPIC_Target.TopLeftCell, pPIC_Target.BottomRightCell)) Is Nothing Then

        If Range(pPIC_Target.TopLeftCell, pPIC_Target.BottomRightCell).Address = _
           Intersect(pRNG_Target, Range(pPIC_Target.TopLeftCell, pPIC_Target.BottomRightCell)).Address Then

            fBLN_IsCheckInRange = True

        End If
    End If

End Function

Private Function fARR_Checknames() As Variant
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: Create array with all possible search queries for checkmarks.
    'Note: DO NOT CHANGE CHECK NAMES, there will be many of these in the environment that could be deleted years from now.

    pSTR_SUBFCN_Name = "fARR_Checknames"
    Dim ARR_Checknames(3) As String

    ARR_Checknames(0) = "*_RedChk_*"
    ARR_Checknames(1) = "*_GreenChk_*"
    ARR_Checknames(2) = "*_BlueChk_*"
    ARR_Checknames(3) = "*_PurpleChk_*"

    fARR_Checknames = ARR_Checknames

End Function

Private Function fBLN_IsPicTickmark() As Boolean
    'Author: Edward Sullivan edwardjohnsullivan@gmail.com
    'Purpose: To deterime if target picture has name that matches the search query array

    pSTR_SUBFCN_Name = "fBLN_IsPicTickmark"
    fBLN_IsPicTickmark = False

    For pINT_Target = LBound(fARR_Checknames) To UBound(fARR_Checknames)

        If pPIC_Target.Name Like fARR_Checknames(pINT_Target) Then

            fBLN_IsPicTickmark = True
            Exit Function

        End If
    Next

End Function

