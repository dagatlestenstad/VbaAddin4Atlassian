VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Settings"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6810
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
    End
End Sub

Private Sub cmdOk_Click()
          
    successfulLogin = True
        
    atlassianURL = Trim(txtAtlassianURL)
    atlassianEmail = Trim(txtAtlassianEmail)
    atlassianToken = Trim(txtAtlassianToken)
        
    lblAtlassianURL.ForeColor = vbBlack
    lblAtlassianEmail.ForeColor = vbBlack
    lblAtlassianToken.ForeColor = vbBlack
       
    If atlassianEmail = vbNullString Then
        MultiPage1.value = 0
        lblAtlassianEmail.ForeColor = vbRed
        txtAtlassianEmail.SetFocus
        Exit Sub
    End If

    If atlassianToken = vbNullString Then
        MultiPage1.value = 0
        lblAtlassianToken.ForeColor = vbRed
        txtAtlassianToken.SetFocus
        Exit Sub
    End If
        
    If atlassianURL = vbNullString Then
        MultiPage1.value = 0
        lblAtlassianURL.ForeColor = vbRed
        txtAtlassianURL.SetFocus
        Exit Sub
    End If
    
    
    Dim Jira As New clsJira
    If Jira.correctCredentianls Then

        If Right(atlassianURL, 1) = "/" Then atlassianURL = Mid(atlassianURL, 1, Len(atlassianURL) - 1)
        
        SaveSetting "VbaAddin4Atlassian", "Settings", "atlassianURL", atlassianURL
        SaveSetting "VbaAddin4Atlassian", "Settings", "atlassianEmail", atlassianEmail
        SaveSetting "VbaAddin4Atlassian", "Settings", "atlassianToken", atlassianToken
        
        Unload Me
    Else
        successfulLogin = False
        MsgBox "Wrong credentials or URL", vbCritical
    End If

End Sub

Private Sub lblLink_Click()
    Call openHyperlink("https://github.com/dagatlestenstad/VbaAddin4Atlassian")
End Sub

Private Sub UserForm_Initialize()

    txtAtlassianURL = GetSetting("VbaAddin4Atlassian", "Settings", "atlassianURL")
    txtAtlassianEmail = GetSetting("VbaAddin4Atlassian", "Settings", "atlassianEmail")
    txtAtlassianToken = GetSetting("VbaAddin4Atlassian", "Settings", "atlassianToken")
    
End Sub
