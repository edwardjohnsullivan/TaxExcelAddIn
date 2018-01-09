VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLS_User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private STR_UserName As String

Private Sub Class_Initalize()
'Author: Edward Sullivan edwardjohnsullivan@gmail.com    
'Purpose: To reinitialize the class "User" if it has not already occoured. 

    If pfBLN_IsCLS_UserInitalized = False Then pSUB_CLS_User_Initialization

End Sub

Public Function pfBLN_IsCLS_UserInitalized() As Boolean
'Author: Edward Sullivan edwardjohnsullivan@gmail.com  
'Purpose: To determine if the initialization of the Class "User" has occoured  

    If IsEmpty(ARR_STR_UserNames) Then

        pfBLN_IsCLS_UserInitalized = True

    Else

        pfBLN_IsCLS_UserInitalized = False

    End If

End Function

Public Property Get STR_FullName() As String
'Author: Edward Sullivan edwardjohnsullivan@gmail.com    
'Purpose: To return the fullname of a user with data from user_data sheet. 

    STR_UserName = Environ$("Username")    ' Windows Username not Office

    If fBLN_UserNameExists(STR_UserName) Then

        If fBLN_UserNameIsValid(STR_UserName) Then

            With Application

                STR_FullName = .Index(ARR_STR_FullNames, .Match(STR_UserName, ARR_STR_UserNames, 0))

            End With

        Else

            STR_FullName = "ERROR"

        End If

    Else

        STR_FullName = "ERROR"

    End If

End Property

Public Property Get STR_CasualName() As String
'Author: Edward Sullivan edwardjohnsullivan@gmail.com    
'Purpose: To return the casual name of a user with data from user_data sheet. 

    STR_UserName = Environ$("Username")    ' Windows Username not Office

    If fBLN_UserNameExists(STR_UserName) Then

        If fBLN_UserNameIsValid(STR_UserName) Then

            With Application

                STR_CasualName = .Index(ARR_STR_CasualNames, .Match(STR_UserName, ARR_STR_UserNames, 0))

            End With

        Else

            STR_CasualName = "ERROR"

        End If

    Else

        STR_CasualName = "ERROR"

    End If

End Property

Public Property Get STR_Initial() As String
'Author: Edward Sullivan edwardjohnsullivan@gmail.com   
'Purpose: To return the initials of a user with data from user_data sheet.  

    STR_UserName = Environ$("Username")    ' Windows Username not Office

    If fBLN_UserNameExists(STR_UserName) Then

        If fBLN_UserNameIsValid(STR_UserName) Then

            With Application

                STR_Initial = .Index(ARR_STR_Initials, .Match(STR_UserName, ARR_STR_UserNames, 0))

            End With

        Else

            STR_Initial = "ERROR"

        End If

    Else

        STR_Initial = "ERROR"

    End If

End Property

Public Property Get STR_OrgLevel() As String
'Author: Edward Sullivan edwardjohnsullivan@gmail.com    
'Purpose: To return the organization level of a user with data from user_data sheet. 

    STR_UserName = Environ$("Username")    ' Windows Username not Office

    If fBLN_UserNameExists(STR_UserName) Then

        If fBLN_UserNameIsValid(STR_UserName) Then

            With Application
                STR_OrgLevel = .Index(ARR_STR_OrgLevels, .Match(STR_UserName, ARR_STR_UserNames, 0))
            End With

        Else

            STR_OrgLevel = "ERROR"

        End If
    Else

        STR_OrgLevel = "ERROR"

    End If

End Property

Public Property Get BLN_IsReviewer() As Boolean
'Author: Edward Sullivan edwardjohnsullivan@gmail.com 
'Purpose: To determine if a user is a reviewer or not.    

    With Application
        If STR_OrgLevel = "Reviewer" Or STR_OrgLevel = "Director" Or STR_OrgLevel = "Trust_Dept" Then
            
            BLN_IsReviewer = True

        Else

            BLN_IsReviewer = False

        End If
    End With

End Property

Private Function fBLN_UserNameExists(STR_UserName As String) As Boolean
'Author: Edward Sullivan edwardjohnsullivan@gmail.com
'Purpose: To determine if a user name can be obtained    

    If STR_UserName = "" Then

        fBLN_UserNameExists = False
        MsgBox "Unable to obtain Username of current user, please retry current action." & vbNewLine & vbNewLine & _
               "If error persists please contact developer."

    Else

        fBLN_UserNameExists = True
        Debug.Print "username exists"
    End If

End Function

Private Function fBLN_UserNameIsValid(STR_UserName As String) As Boolean
'Author: Edward Sullivan edwardjohnsullivan@gmail.com
'Purpose: To determine if a user name exists in the UserName array which is populated from the user_data worksheet.     

    With Application
        If IsError(.Match(STR_UserName, ARR_STR_UserNames, 0)) Then

            fBLN_UserNameIsValid = False
            MsgBox "Current user not set up for use of the KTC Tax Excel Add-in." & vbNewLine & vbNewLine & _
                   "Please contact developer to set up new user"

        Else

            fBLN_UserNameIsValid = True
            Debug.Print "username is valid"

        End If
    End With

End Function


