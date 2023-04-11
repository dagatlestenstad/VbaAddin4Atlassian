VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJira"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private RestClient As New clsRestClient

Private Sub Class_initialize()
    If Not successfulLogin Then
        If correctCredentianls Then successfulLogin = True
    End If
End Sub

Private Sub Class_Terminate()
    Set RestClient = Nothing
End Sub

Public Sub saveIssueAttachment(filename As String, file As String, jiraKey As String)

    Dim data() As Byte
    
    data = StrConv("--" & boundary & vbCrLf & _
        "Content-Disposition: form-data; name=""file""; filename=""" & filename & """" & vbCrLf & _
        "Content-Type: application/octet-stream" & vbCrLf & vbCrLf & StrConv(file, vbUnicode) & vbCrLf & _
        "--" & boundary & "--", vbFromUnicode)
        
    Call RestClient.Service("POST", "/rest/api/" & atlassianRestVersion & "/issue/" & jiraKey & "/attachments", data)
     
End Sub

Public Function correctCredentianls() As Boolean
    
    Dim rest As restResponse
    
    rest = RestClient.Service("GET", "/rest/api/" & atlassianRestVersion & "/myself")
    
    If rest.Status = 200 Then correctCredentianls = True

End Function

Public Function getProject() As Collection
    
    Dim rest As restResponse
    
    Dim json As Object
    Dim item As Variant
    
    Dim project As clsProject
    Dim projects As New Collection
    
    rest = RestClient.Service("GET", "/rest/api/" & atlassianRestVersion & "/project")
    
    Set json = JsonConverter.ParseJson(rest.Body)
    
    For Each item In json
        Set project = New clsProject
        
        project.id = item("id")
        project.key = item("key")
        project.Name = item("name")
        
        projects.Add project
    Next
    
    Set getProject = projects

End Function

Public Function getIssueTypes(jiraProjectID As String) As Collection
    
    Dim rest As restResponse
    
    Dim json As Object
    Dim item As Variant
    
    Dim issueType As clsIssuetype
    Dim issueTypes As New Collection
        
    rest = RestClient.Service("GET", "/rest/api/" & atlassianRestVersion & "/project/" & jiraProjectID & "/statuses")
    
    Set json = JsonConverter.ParseJson(rest.Body)
    
     For Each item In json
        Set issueType = New clsIssuetype
            
        issueType.id = item("id")
        issueType.Name = item("name")
        issueType.subtask = item("subtask")
        
        issueTypes.Add issueType
    Next
    
    Set getIssueTypes = issueTypes
    
End Function

Public Function createIssue(project As String, issueType As String, summary As String, description As String) As String

    Dim rest As restResponse
    Dim data As String

    data = " { ""fields"" : { " & _
        """project"": { ""key"": """ & project & """ }, "

        data = data & """summary"": """ & summary & """, "
        
        If Not Trim(description) = vbNullString Then data = data & """description"": { ""content"": [ { ""content"": [ { ""text"": " & ConvertToJson(description) & ", ""type"": ""text"" } ], ""type"": ""paragraph"" } ], ""type"": ""doc"", ""version"": 1 }, "
           
        data = data & """issuetype"" : { ""name"" : """ & issueType & """ } } }"
        
        rest = RestClient.Service("POST", "/rest/api/" & atlassianRestVersion & "/issue/", data)
        
        If rest.Status = 201 Then
            Dim json As Object
            Set json = JsonConverter.ParseJson(rest.Body)
            createIssue = json("key")
        Else
            MsgBox rest.Body, vbCritical, rest.Status
        End If
       
End Function
