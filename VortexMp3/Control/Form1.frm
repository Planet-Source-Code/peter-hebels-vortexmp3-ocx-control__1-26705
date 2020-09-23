VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About VortexSDPlay"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Line Line6 
      X1              =   4800
      X2              =   4920
      Y1              =   960
      Y2              =   840
   End
   Begin VB.Line Line5 
      X1              =   4800
      X2              =   4920
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line4 
      X1              =   5040
      X2              =   4800
      Y1              =   720
      Y2              =   840
   End
   Begin VB.Line Line3 
      X1              =   4680
      X2              =   5040
      Y1              =   600
      Y2              =   720
   End
   Begin VB.Line Line2 
      X1              =   5160
      X2              =   4680
      Y1              =   480
      Y2              =   600
   End
   Begin VB.Line Line1 
      X1              =   4560
      X2              =   5160
      Y1              =   360
      Y2              =   480
   End
   Begin VB.Label Label4 
      Caption         =   "I'am not responsible for any damages. "
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "You are free to use this control in your app's"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Control made by Peter Hebels."
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "VortexSDPlay Control, for playing wave and mpeg files."
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Form1.frx":0000
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

