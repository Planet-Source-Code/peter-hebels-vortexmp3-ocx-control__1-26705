VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{E920B5F6-807D-11D5-B73C-8886ED076B33}#1.0#0"; "VORTEXSD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wave Play Mp3-Player"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Tag"
      Height          =   2295
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   4815
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1800
         TabIndex        =   15
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         TabIndex        =   14
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1800
         TabIndex        =   13
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Top             =   1680
         Width           =   2895
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Load Tag"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1800
         TabIndex        =   21
         Top             =   1920
         Width           =   2895
      End
      Begin VortexSDPlay.VortexSDCtrl VortexSDCtrl1 
         Left            =   1080
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         SongName        =   ""
      End
      Begin VB.Label Label5 
         Caption         =   "Year:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Title:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Comment:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Album:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Artist:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control"
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton Command7 
         Caption         =   "About"
         Height          =   615
         Left            =   3840
         TabIndex        =   39
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Load File"
         Height          =   615
         Left            =   3000
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Resume"
         Height          =   615
         Left            =   1920
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Pause"
         Height          =   615
         Left            =   1320
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Stop"
         Height          =   615
         Left            =   720
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Play"
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Wave-Out"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   4815
      Begin ComctlLib.ProgressBar ProgressBar10 
         Height          =   255
         Left            =   1200
         TabIndex        =   23
         Top             =   2400
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar ProgressBar9 
         Height          =   255
         Left            =   1200
         TabIndex        =   22
         Top             =   2160
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar ProgressBar8 
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Top             =   1920
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar ProgressBar7 
         Height          =   255
         Left            =   1200
         TabIndex        =   19
         Top             =   1680
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar ProgressBar6 
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         Top             =   1440
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar ProgressBar5 
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   1200
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar ProgressBar4 
         Height          =   255
         Left            =   1200
         TabIndex        =   1
         Top             =   960
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
         Max             =   1000
      End
      Begin ComctlLib.ProgressBar ProgressBar3 
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   720
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
         Max             =   5000
      End
      Begin ComctlLib.ProgressBar ProgressBar2 
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   480
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
         Max             =   5000
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
         Max             =   50000
      End
      Begin VB.Label Label15 
         Caption         =   "31Hz"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "62Hz"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "125Hz"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "250Hz"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "500Hz"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "1kHz"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "3kHz"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "6kHz"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "12kHz"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "16kHz"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3600
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   4560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'Before you start, first load the VortexSD-control
'after this you can place the VortexSD-Control on the
'form "You can load it by pressing CTRL+T then press
'the Browse button and select the file from the same folder
'where the project files are.
'**********************************************************

Private Sub Command1_Click()
VortexSDCtrl1.CommandMe = "Stop"    'Send command STOP
End Sub

Private Sub Command2_Click()
VortexSDCtrl1.CommandMe = "Play"    'Send command PLAY
End Sub

Private Sub Command3_Click()
VortexSDCtrl1.CommandMe = "Pause"   'Send command PAUSE
End Sub

Private Sub Command4_Click()
VortexSDCtrl1.CommandMe = "UnPause" 'Send command UNPAUSE
End Sub

Private Sub Command7_Click()
VortexSDCtrl1.CommandMe = "AboutBox" 'Send command ABOUTBOX
End Sub


Private Sub Command5_Click()
Text1.Text = VortexSDCtrl1.Album    'Import ALBUM string
Text2.Text = VortexSDCtrl1.Artist   'Import ARTIST string
Text3.Text = VortexSDCtrl1.Comment  'Import COMMENT string
Text4.Text = VortexSDCtrl1.Title    'Import TITLE string
Text5.Text = VortexSDCtrl1.Year     'Import YEAR string
End Sub

Private Sub Command6_Click()
CommonDialog1.DialogTitle = "Open"  'Show open dialog
CommonDialog1.ShowOpen
VortexSDCtrl1.SongName = CommonDialog1.FileName   'Send selected songname to Control
End Sub

Private Sub Form_Load()
VortexSDCtrl1.CommandMe = "StartVis"   'Start WAVE Visualization
'VortexSDCtrl1.CommandMe = "StartVolCtrl"  'Set Volume

End Sub

Private Sub Form_Unload(Cancel As Integer)
VortexSDCtrl1.CommandMe = "StopVis"     'Stop WAVE Visualization output
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Max = VortexSDCtrl1.Freq16kHzMaxOut   'Visualization output strings maximum
ProgressBar2.Max = VortexSDCtrl1.Freq12kHzMaxOut
ProgressBar3.Max = VortexSDCtrl1.Freq6kHzMaxOut
ProgressBar4.Max = VortexSDCtrl1.Freq3kHzMaxOut
ProgressBar5.Max = VortexSDCtrl1.Freq1kHzMaxOut
ProgressBar6.Max = VortexSDCtrl1.Freq500HzMaxOut
ProgressBar7.Max = VortexSDCtrl1.Freq250HzMaxOut
ProgressBar8.Max = VortexSDCtrl1.Freq125HzMaxOut
ProgressBar9.Max = VortexSDCtrl1.Freq62HzMaxOut
ProgressBar10.Max = VortexSDCtrl1.Freq31HzMaxOut


ProgressBar1.Min = VortexSDCtrl1.Freq16kHzMinOut   'Visualization output string minimum
ProgressBar2.Min = VortexSDCtrl1.Freq12kHzMinOut
ProgressBar3.Min = VortexSDCtrl1.Freq6kHzMinOut
ProgressBar4.Min = VortexSDCtrl1.Freq3kHzMinOut
ProgressBar5.Min = VortexSDCtrl1.Freq1kHzMinOut
ProgressBar6.Min = VortexSDCtrl1.Freq500HzMinOut
ProgressBar7.Min = VortexSDCtrl1.Freq250HzMinOut
ProgressBar8.Min = VortexSDCtrl1.Freq125HzMinOut
ProgressBar9.Min = VortexSDCtrl1.Freq62HzMinOut
ProgressBar10.Min = VortexSDCtrl1.Freq31HzMinOut


ProgressBar1.Value = VortexSDCtrl1.Freq16kHzOut    'Output from control to progressbar
ProgressBar2.Value = VortexSDCtrl1.Freq12kHzOut
ProgressBar3.Value = VortexSDCtrl1.Freq6kHzOut
ProgressBar4.Value = VortexSDCtrl1.Freq3kHzOut
ProgressBar5.Value = VortexSDCtrl1.Freq1kHzOut
ProgressBar6.Value = VortexSDCtrl1.Freq500HzOut
ProgressBar7.Value = VortexSDCtrl1.Freq250HzOut
ProgressBar8.Value = VortexSDCtrl1.Freq125HzOut
ProgressBar9.Value = VortexSDCtrl1.Freq62HzOut
ProgressBar10.Value = VortexSDCtrl1.Freq31HzOut

End Sub



