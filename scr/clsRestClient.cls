VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRestClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Class_initialize()

    If Not successfulLogin Then
        boundary = createBoundary
        
        If atlassianURL = vbNullString Or atlassianEmail = vbNullString Or atlassianToken = vbNullString Then
            atlassianURL = GetSetting("VbaAddin4Atlassian", "Settings", "atlassianURL")
            atlassianEmail = GetSetting("VbaAddin4Atlassian", "Settings", "atlassianEmail")
            atlassianToken = GetSetting("VbaAddin4Atlassian", "Settings", "atlassianToken")
        End If
          
        If atlassianURL = vbNullString Or atlassianEmail = vbNullString Or atlassianToken = vbNullString Then
            frmSettings.Show
        End If
    End If
    
End Sub

Public Function Service(methods As String, url As String, Optional data As Variant) As restResponse

    On Error GoTo ErrHandler
           
    Dim rest As restResponse
     
    Dim RestClient As Object
    Set RestClient = CreateObject("WinHttp.WinHttpRequest.5.1")

     With RestClient
        .Open methods, atlassianURL & url, False
        .setRequestHeader "Authorization", "Basic " & EncodeBase64(atlassianEmail & ":" & atlassianToken)
    
        If Not VarType(data) = 8209 Then
            .setRequestHeader "Accept", "application/json"
            .setRequestHeader "Content-Type", "application/json"
        Else
            .setRequestHeader "X-Atlassian-Token", "nocheck"
            .setRequestHeader "Content-Type", "multipart/form-data; charset=""ISO-8859-1""; boundary=" & boundary
        End If
        
        .Send data
        
        rest.Status = .Status
        rest.Body = .responseText
        Service = rest
    End With
        
    Set RestClient = Nothing
    
    Exit Function
    

ErrHandler:

    MsgBox err.description, vbCritical, err.Number
    
End Function

Private Function EncodeBase64(text As String) As String

    Dim arrData() As Byte
    arrData = StrConv(text, vbFromUnicode)
  
    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement
  
    Set objXML = New MSXML2.DOMDocument
    Set objNode = objXML.createElement("b64")
  
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = Replace(objNode.text, vbLf, "")
  
    Set objNode = Nothing
    Set objXML = Nothing
    
End Function

Private Function createBoundary() As String
    
    Dim characters As String
    Dim boundary As String
    Dim i As Integer
    
    characters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    boundary = ""
        
    For i = 1 To 20
        boundary = boundary & Mid(characters, Int(Len(characters) * Rnd + 1), 1)
    Next
    
    createBoundary = boundary

End Function
