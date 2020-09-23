VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TagEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private blnTagExists As Boolean
Private strTitle As String * 30
Private strArtist As String * 30
Private strAlbum As String * 30
Private strYear As String * 4
Private strComment As String * 30
Private strFilename As String

Public Function LoadMP3(ByVal Filename As String) As Boolean
  Dim FF As Integer
  Dim strTag As String * 128
  
  Call ClearProperties
  strFilename = Filename
  
  On Error GoTo LoadMp3Tag_Error
  
  If FileLen(Filename) > 128 Then
    FF = FreeFile
    Open Filename For Binary As FF
      Get FF, FileLen(Filename) - 127, strTag
      
      If Left$(strTag, 3) = "TAG" Then
              
        blnTagExists = True
        strTitle = Mid$(strTag, 4, 30)
        strArtist = Mid$(strTag, 34, 30)
        strAlbum = Mid$(strTag, 64, 30)
        strYear = Mid$(strTag, 94, 4)
        strComment = Mid$(strTag, 98, 30)
        
      End If
    Close FF
  End If

  LoadMP3 = True
  Exit Function

LoadMp3Tag_Error:
  LoadMP3 = False
  Exit Function

End Function

Public Property Let Filename(ByVal Value As String)
  strFilename = Value
End Property

Public Property Get Filename() As String
  Filename = strFilename
End Property

Public Property Let Comment(ByVal Value As String)
  strComment = Value
End Property

Public Property Get Comment() As String
  Comment = Trim$(strComment)
End Property

Public Property Let Year(ByVal Value As String)
  strYear = Value
End Property

Public Property Get Year() As String
  Year = Trim$(strYear)
End Property

Public Property Let Album(ByVal Value As String)
  strAlbum = Value
End Property

Public Property Get Album() As String
  Album = Trim$(strAlbum)
End Property

Public Property Let Artist(ByVal Value As String)
  strArtist = Value
End Property

Public Property Get Artist() As String
  Artist = Trim$(strArtist)
End Property

Public Property Let Title(ByVal Value As String)
  strTitle = Value
End Property

Public Property Get Title() As String
  Title = Trim$(strTitle)
End Property

Public Property Let TagExists(ByVal Value As Boolean)
  blnTagExists = Value
End Property

Public Property Get TagExists() As Boolean
  TagExists = blnTagExists
End Property

Private Sub ClearProperties()
  blnTagExists = False
  strTitle = ""
  strArtist = ""
  strAlbum = ""
  strYear = ""
  strComment = ""
End Sub

Private Function SpaceOut(strText As String, intLength As Integer) As String
    If Len(strText) > intLength Then
        SpaceOut = Left(strText, intLength)
        Exit Function
    Else
        SpaceOut = strText + String(intLength - Len(strText), " ")
    End If
End Function



