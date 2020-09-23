Attribute VB_Name = "Module1"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Dim HeaderPosition As String

Function GetHeader(Filename As String)
   On Error GoTo noFileDe
   Dim X As Byte
   Dim kbps As Integer
   Dim khz As Single
   Dim HeaderPosition As Long
   Dim FrameLength As Long
   Dim nFrames As Long
   Dim Xing As Boolean
  
   fn = FreeFile
   Open Filename For Binary As #fn
   
   Get #fn, 1, X
   If X <> 255 Then
      If X <> 73 Then Exit Function
   End If
   HeaderPosition = 1
   Get #fn, 2, X
   If (X < 250 Or X > 251) Then
      If X = 68 Then
         Get #fn, 3, X
         If X = 51 Then
            Get #fn, 7, X
            d = CLng(X) * 20917152
            Get #fn, 8, X
            d = CLng(d) + (CLng(X) * 16384)
            Get #fn, 9, X
            d = CLng(d) + (CLng(X) * 128)
            Get #fn, 10, X
            d = d + X
            If d > LOF(fn) Or d > 2147483647 Then Exit Function
            HeaderPosition = d + 11
         End If
      Else
         CheckMp3 = False
         Close #fn
         Exit Function
      End If
   End If
   
  Close #fn

If HeaderPosition <= 2000 Then GoTo noFileDe
MsgBox "Unable to Get Header", vbCritical, "Error"

Exit Function
noFileDe:
Close #fn
End Function

