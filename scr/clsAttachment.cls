VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAttachment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As String
Private pFilename As String
Private pData As String

Public Property Get id() As String
    id = pId
End Property

Public Property Let id(value As String)
    pId = value
End Property

Public Property Get filename() As String
    filename = pFilename
End Property

Public Property Let filename(value As String)
    pFilename = value
End Property

Public Property Get data() As String
    data = pData
End Property

Public Property Let data(value As String)
    pData = value
End Property
