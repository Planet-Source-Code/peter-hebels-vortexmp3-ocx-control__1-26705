VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MP3Genra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sName As String
Private iID As Integer

Public Property Get Name() As String
  Name = sName
End Property

Public Property Get ID() As Integer
  ID = iID
End Property

Friend Sub Constructor(Name As String, ID As Integer)
  sName = Name
  iID = ID
End Sub
