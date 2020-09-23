VERSION 5.00
Begin VB.UserControl VortexSDCtrl 
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "UserControl1.ctx":0000
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "UserControl1.ctx":030A
   Windowless      =   -1  'True
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3240
      Top             =   1560
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3000
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "VortexSDCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************************************
'VortexSD Mp3Play Control project by Peter Hebels Website "www.grworld.com/megagsite/peterspagina.html    *
'Iam not responsible for any damages may caused by this project                                           *
'**********************************************************************************************************


Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Dim ShortPathAndFie As String
Dim ShortPath As Long
Dim tmp As String * 255
Dim s As String * 30
Dim VolMaximum As String
Dim Freq500HzMax As Integer
Dim Freq500HzMin As Integer
Dim Freq250HzMax As Integer
Dim Freq250HzMin As Integer
Dim Freq125HzMax As Integer
Dim Freq125HzMin As Integer
Dim Freq62HzMax As Integer
Dim Freq62HzMin As Integer
Dim Freq31HzMax As Integer
Dim Freq31HzMin As Integer
Dim Freq1kHzMax As Integer
Dim Freq1kHzMin As Integer
Dim Freq3kHzMax As Integer
Dim Freq3kHzMin As Integer
Dim Freq6kHzMax As Integer
Dim Freq6kHzMin As Integer
Dim Freq12kHzMax As Integer
Dim Freq12kHzMin As Integer
Dim Freq16kHzMax As Integer
Dim Freq16kHzMin As Integer

Dim txtTitle As String
Dim txtArtist As String
Dim txtAlbum As String
Dim txtYear As String
Dim txtComment As String
Dim oTagEditor As New TagEditor
'-------------------------------------------------------------------
Dim hmixer As Long
Dim inputVolCtrl As MIXERCONTROL
Dim outputVolCtrl As MIXERCONTROL
Dim volCtrl As MIXERCONTROL
Dim rc As Long
Dim ok As Boolean
Dim mxcd As MIXERCONTROLDETAILS
Dim vol As MIXERCONTROLDETAILS_SIGNED
Dim volume As Long
Dim volHmem As Long
Private VU As VULights
Private FreqNum As Frequency
'-------------------------------------------------------------------

Private Sub VolVal(VolIs As Long, VolFreq As Double)
For FreqNum = 0 To 9
Next FreqNum
VolIs = volume * 327.67
VolFreq = VU.Freq(FreqNum)
VU.FreqVal = VolIs * VolFreq
End Sub

Public Property Get SongName() As String
   SongName = Label1.Caption
End Property

Public Property Let SongName(NewSongName As String)
   sndPlaySound None, 4
   mciSendString "close mpeg", 0, 0, 0
   Label1.Caption = NewSongName
   
   If SongName = "" Then Exit Property
   
   LoadTag
   PropertyChanged "SongName"
End Property


Private Sub Timer2_Timer()
    VU.VolLev = volume / 327.67
    If (volume < 0) Then volume = -volume
    If (1 = 1) Then
    mxcd.dwControlID = outputVolCtrl.dwControlID
    mxcd.item = outputVolCtrl.cMultipleItems
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
    If (volume < 0) Then volume = -volume
    End If

End Sub

Private Sub UserControl_Initialize()
Label1.Caption = ""

End Sub

Private Sub UserControl_Resize()
 UserControl.Width = 480
 UserControl.Height = 480
End Sub

Private Sub UserControl_Terminate()
sndPlaySound None, 4
mciSendString "close mpeg", 0, 0, 0

If (fRecording = True) Then
    StopInput
End If
GlobalFree volHmem
End Sub

Public Property Let CommandMe(Commando As String)
   If Commando = "Stop" Then GoTo StopPlay
   If Commando = "Play" Then GoTo BeginPlay
   If Commando = "Pause" Then GoTo PausePlay
   If Commando = "UnPause" Then GoTo UnPauseMpegFile
   If Commando = "AboutBox" Then GoTo ShowAbout
   If Commando = "StartVis" Then GoTo StartInputVis
   If Commando = "StopVis" Then GoTo StopInputVis
   If Commando = "StartVolCtrl" Then GoTo SvCtrl
   If Commando = "AllStop" Then GoTo StopAllFnc
   
   MsgBox "Command Not Accepted. useable Commands are: Stop, Play, Pause, UnPause, StartVis, StopVis, StartVolCtrl, AllStop, AboutBox. "
   Exit Property

StopPlay:
   sndPlaySound None, 4
   mciSendString "close mpeg", 0, 0, 0
     
   Exit Property

BeginPlay:
   If Right(Label1.Caption, 3) = "wav" Then GoTo PlayWaveFile
   If Right(Label1.Caption, 3) = "mp3" Then GoTo PlayMpegFile
   MsgBox "Cannot Decode This Media", vbCritical
   Exit Property

PausePlay:
   If Right(Label1.Caption, 3) = "mp3" Then GoTo PauseMpegFile
   Exit Property

PlayWaveFile:
   sndPlaySound Label1.Caption, 1
   
   Exit Property

PlayMpegFile:
   GetHeader Label1.Caption
   
   ShortPath = GetShortPathName(Label1.Caption, tmp, 255)
   ShortPathAndFie = Left$(tmp, ShortPath)
   mciSendString "close mpeg", 0, 0, 0
   mciSendString "open " & ShortPathAndFie & " type MPEGVideo Alias mpeg", 0&, 0&, 0&
   mciSendString "play mpeg", 0, 0, 0
      
   LoadTag
   Exit Property
   
PauseMpegFile:
   mciSendString "pause mpeg", 0, 0, 0
   
   Exit Property

UnPauseMpegFile:
   mciSendString "play mpeg", 0, 0, 0
   
   Exit Property

ShowAbout:
   Form1.Show

   Exit Property
   
StartInputVis:
   Timer2.Enabled = True
   Timer2.Interval = 3

   rc = mixerOpen(hmixer, DEVICEID, 0, 0, 0)
   If ((MMSYSERR_NOERROR <> rc)) Then
   Exit Property
   End If
   
   ok = GetControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_PEAKMETER, outputVolCtrl)
   If (ok = True) Then
    Freq500HzMax = Frequency.Freq500Hz + 1
    Freq500HzMin = Frequency.Freq500Hz
    Freq250HzMax = Frequency.Freq250Hz + 1
    Freq250HzMin = Frequency.Freq250Hz
    Freq125HzMax = Frequency.Freq125Hz + 1
    Freq125HzMin = Frequency.Freq125Hz
    Freq62HzMax = Frequency.Freq62Hz + 1
    Freq62HzMin = Frequency.Freq62Hz
    Freq31HzMax = Frequency.Freq31Hz + 1
    Freq31HzMin = Frequency.Freq31Hz
    Freq1kHzMax = Frequency.Freq1kHz + 1
    Freq1kHzMin = Frequency.Freq1kHz
    Freq3kHzMax = Frequency.Freq3kHz + 1
    Freq3kHzMin = Frequency.Freq3kHz
    Freq6kHzMax = Frequency.Freq6kHz + 1
    Freq6kHzMin = Frequency.Freq6kHz
    Freq12kHzMax = Frequency.Freq12kHz + 1
    Freq12kHzMin = Frequency.Freq12kHz
    Freq16kHzMax = Frequency.Freq16kHz + 1
    Freq16kHzMin = Frequency.Freq16kHz
   Else
      MsgBox "waveout meter not supported by soundcard"
   End If
   
   mxcd.cbStruct = Len(mxcd)
   volHmem = GlobalAlloc(&H0, Len(volume))
   mxcd.paDetails = GlobalLock(volHmem)
   mxcd.cbDetails = Len(volume)
   mxcd.cChannels = 1
   
   Exit Property
   
StopInputVis:
If (fRecording = True) Then
    StopInput
End If
GlobalFree volHmem
  
   Exit Property

SvCtrl:
         rc = mixerOpen(hmixer, 0, 0, 0, 0)
         If ((MMSYSERR_NOERROR <> rc)) Then
             MsgBox "Couldn't open the mixer."
             Exit Property
             End If

         ok = GetVolumeControl(hmixer, _
                              MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                              MIXERCONTROL_CONTROLTYPE_VOLUME, _
                              volCtrl)
         If (ok = True) Then
             
             VolMaximum = volCtrl.lMaximum
         End If
   Exit Property
   
StopAllFnc:

sndPlaySound None, 4
mciSendString "close mpeg", 0, 0, 0

If (fRecording = True) Then
    StopInput
End If
GlobalFree volHmem

   Exit Property

End Property

Public Property Get Artist() As String
   Artist = txtArtist
End Property

Public Property Get Title() As String
   Title = txtTitle
End Property

Public Property Get Year() As String
   Year = txtYear
End Property

Public Property Get Album() As String
   Album = txtAlbum
End Property

Public Property Get Comment() As String
   Comment = txtComment
End Property

Public Function LoadTag()
  Dim sMsg As String
  Dim lLoop As Long
    
  Call oTagEditor.LoadMP3(Label1.Caption)
  
  txtTitle = oTagEditor.Title
  txtArtist = oTagEditor.Artist
  txtAlbum = oTagEditor.Album
  txtYear = oTagEditor.Year
  txtComment = oTagEditor.Comment
  
End Function

Public Property Get Freq500HzOut() As String
FreqNum = Freq500Hz
For VU.InOutLev = CDbl(VU.VolLev * 0.2) To FreqNum
Next VU.InOutLev
   Freq500HzOut = VU.InOutLev
End Property

Public Property Get Freq250HzOut() As String
FreqNum = Freq250Hz
For VU.InOutLev = CDbl(VU.VolLev * 0.4) To FreqNum
Next VU.InOutLev
   Freq250HzOut = VU.InOutLev
End Property

Public Property Get Freq125HzOut() As String
FreqNum = Freq125Hz
For VU.InOutLev = CDbl(VU.VolLev * 0.8) To FreqNum
Next VU.InOutLev
   Freq125HzOut = VU.InOutLev
End Property

Public Property Get Freq62HzOut() As String
FreqNum = Freq62Hz
For VU.InOutLev = CDbl(VU.VolLev * 1.61290322580645E-02) To FreqNum
Next VU.InOutLev
   Freq62HzOut = VU.InOutLev
End Property

Public Property Get Freq31HzOut() As String
FreqNum = Freq31Hz
For VU.InOutLev = CDbl(VU.VolLev * 0.032258064516129) To FreqNum
Next VU.InOutLev
   Freq31HzOut = VU.InOutLev
End Property

Public Property Get Freq1kHzOut() As String
FreqNum = Freq1kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.01) To FreqNum
Next VU.InOutLev
   Freq1kHzOut = VU.InOutLev
End Property

Public Property Get Freq3kHzOut() As String
FreqNum = Freq3kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.03) To FreqNum
Next VU.InOutLev
   Freq3kHzOut = VU.InOutLev
End Property

Public Property Get Freq6kHzOut() As String
FreqNum = Freq6kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.06) To FreqNum
Next VU.InOutLev
   Freq6kHzOut = VU.InOutLev
End Property

Public Property Get Freq12kHzOut() As String
FreqNum = Freq12kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.12) To FreqNum
Next VU.InOutLev
   Freq12kHzOut = VU.InOutLev
End Property

Public Property Get Freq16kHzOut() As String
FreqNum = Freq16kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.16) To FreqNum
Next VU.InOutLev
   Freq16kHzOut = VU.InOutLev
End Property

'----------------------------------------------------------------------
Public Property Get Freq500HzMaxOut() As String
   Freq500HzMaxOut = Freq500HzMax
End Property

Public Property Get Freq250HzMaxOut() As String
   Freq250HzMaxOut = Freq250HzMax
End Property

Public Property Get Freq125HzMaxOut() As String
   Freq125HzMaxOut = Freq125HzMax
End Property

Public Property Get Freq62HzMaxOut() As String
   Freq62HzMaxOut = Freq62HzMax
End Property

Public Property Get Freq31HzMaxOut() As String
   Freq31HzMaxOut = Freq31HzMax
End Property

Public Property Get Freq1kHzMaxOut() As String
   Freq1kHzMaxOut = Freq1kHzMax
End Property

Public Property Get Freq3kHzMaxOut() As String
   Freq3kHzMaxOut = Freq3kHzMax
End Property

Public Property Get Freq6kHzMaxOut() As String
   Freq6kHzMaxOut = Freq6kHzMax
End Property

Public Property Get Freq12kHzMaxOut() As String
   Freq12kHzMaxOut = Freq12kHzMax
End Property

Public Property Get Freq16kHzMaxOut() As String
   Freq16kHzMaxOut = Freq16kHzMax
End Property
'------------------------------------------------------------------------

Public Property Get Freq500HzMinOut() As String
   Freq500HzMinOut = Freq500HzMin
End Property

Public Property Get Freq250HzMinOut() As String
   Freq250HzMinOut = Freq250HzMin
End Property

Public Property Get Freq125HzMinOut() As String
   Freq125HzMinOut = Freq125HzMin
End Property

Public Property Get Freq62HzMinOut() As String
   Freq62HzMinOut = Freq62HzMin
End Property

Public Property Get Freq31HzMinOut() As String
   Freq31HzMinOut = Freq31HzMin
End Property

Public Property Get Freq1kHzMinOut() As String
   Freq1kHzMinOut = Freq1kHzMin
End Property

Public Property Get Freq3kHzMinOut() As String
   Freq3kHzMinOut = Freq3kHzMin
End Property

Public Property Get Freq6kHzMinOut() As String
   Freq6kHzMinOut = Freq6kHzMin
End Property

Public Property Get Freq12kHzMinOut() As String
   Freq12kHzMinOut = Freq12kHzMin
End Property

Public Property Get Freq16kHzMinOut() As String
   Freq16kHzMinOut = Freq16kHzMin
End Property

Public Property Get VolMax() As String
   VolMax = VolMaximum
End Property

Public Property Let VolumeSet(NewVolumeSet As String)
   volumo = CLng(NewVolumeSet)
   If NewVolumeSet <= VolMaximum Then GoTo VolErr
   SetVolumeControl hmixer, volCtrl, volumo
   
Exit Property
VolErr:
MsgBox "Volume set too high", vbCritical, "Error"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    SongName = PropBag.ReadProperty("SongName", "")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "SongName", SongName
End Sub


