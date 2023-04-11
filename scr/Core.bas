Attribute VB_Name = "Core"
Public Const atlassianRestVersion As Integer = 3

Public atlassianURL As String
Public atlassianEmail As String
Public atlassianToken As String

Public successfulLogin  As Boolean
Public boundary As String

Public Type restResponse
    Status As Integer
    Body As String
End Type

'Open url in default browser
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub openHyperlink(url)
    ShellExecute 0, vbNullString, url, vbNullString, vbNullString, vbNormalFocus
End Sub

Public Sub createJiraIssue()
    frmCreateJiraIssue.Show
End Sub

Public Sub openSettings()
    frmSettings.Show
End Sub

Public Function ReadFile(ByVal sFilepath As String) As String
    Dim oStream As Object
    Set oStream = CreateObject("ADODB.Stream")
    
    With oStream
        .Type = 1
        .Open
        .LoadFromFile sFilepath
        ReadFile = .Read
        .Close
    End With
    
    Set oStream = Nothing
End Function
