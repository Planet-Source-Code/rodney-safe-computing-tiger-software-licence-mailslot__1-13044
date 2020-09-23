VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4605
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   330
      Left            =   3645
      TabIndex        =   1
      Top             =   1080
      Width           =   690
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1710
      Top             =   675
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   -45
      Picture         =   "frmInfo.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1020
   End
   Begin VB.Label info 
      Caption         =   "This program is allready running"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   900
      TabIndex        =   0
      Top             =   135
      Width           =   3435
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload frmMain
End Sub

Private Sub Timer1_Timer()
Static I As Integer

Select Case I
Case 2, 3
        info = "Rodney Safe Computing"
        info.Alignment = 2
Case 4, 5
        info = "Shutdown....."
        info.Alignment = 2
Case 6
        Unload Me
End Select
I = I + 1
End Sub
