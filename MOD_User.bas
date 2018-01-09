Attribute VB_Name = "MOD_User"
Option Explicit
'Following vars need to be public as they are used in the Class CLS_User

Public ARR_STR_UserNames() As String
Public ARR_STR_FullNames() As String
Public ARR_STR_CasualNames() As String
Public ARR_STR_Initials() As String
Public ARR_STR_OrgLevels() As String

Public Sub pSUB_CLS_User_Initialization()
'Author: Edward Sullivan edwardjohnsullivan@gmail.com    
'Purpose: To initilize certain arrays for use in the Class "User"
'Note: We are not using the Class Initialize because we only want to open this workbook once when the add-in loads. 

    Dim WB_UserData As Workbook
    Set WB_UserData = Workbooks.Open(Filename:="J:\TAX\KTC Tax Excel Add-in\External_Variables\USER_DATA.xlsx", ReadOnly:=True)

    Dim WS_UserData As Worksheet
    Set WS_UserData = WB_UserData.Sheets("USER_DATA")

    With WS_UserData
        ARR_STR_UserNames = pfARR_RangetoStringArray(.Range("Windows_User_Name"))
        ARR_STR_FullNames = pfARR_RangetoStringArray(.Range("Full_Name"))
        ARR_STR_CasualNames = pfARR_RangetoStringArray(.Range("Casual_Name"))
        ARR_STR_Initials = pfARR_RangetoStringArray(.Range("Initials"))
        ARR_STR_OrgLevels = pfARR_RangetoStringArray(.Range("Organizational_Level"))
    End With

    WB_UserData.Close

End Sub

