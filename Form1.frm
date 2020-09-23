VERSION 5.00
Object = "{63D7CAFE-E32B-4296-9D18-899D4C4DA5FF}#3.0#0"; "Licence.ocx"
Begin VB.Form frmMain 
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8895
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMaxUserLicence 
      Height          =   285
      Left            =   4680
      TabIndex        =   24
      Text            =   "1"
      Top             =   2925
      Width           =   510
   End
   Begin VB.ListBox lstUsers 
      Height          =   1230
      Left            =   6165
      TabIndex        =   22
      Top             =   3465
      Width           =   2625
   End
   Begin VB.OptionButton options 
      Caption         =   "Use deffault Settings"
      Height          =   240
      Index           =   0
      Left            =   1350
      TabIndex        =   20
      Top             =   1755
      Value           =   -1  'True
      Width           =   5550
   End
   Begin VB.Frame Frame1 
      Height          =   3030
      Left            =   0
      TabIndex        =   9
      Top             =   -135
      Width           =   8880
      Begin VB.OptionButton options 
         Caption         =   "Use Above defined user settings "
         Height          =   240
         Index           =   1
         Left            =   1350
         TabIndex        =   19
         Top             =   2160
         Width           =   5550
      End
      Begin VB.OptionButton options 
         Caption         =   "Create Mailslot server licence based on application title"
         Height          =   240
         Index           =   2
         Left            =   1350
         TabIndex        =   18
         Top             =   2430
         Width           =   5550
      End
      Begin VB.TextBox txtClientSlotname 
         Height          =   330
         Left            =   1350
         TabIndex        =   13
         Top             =   270
         Width           =   5550
      End
      Begin VB.TextBox txtServerSlotName 
         Height          =   330
         Left            =   1350
         TabIndex        =   12
         Top             =   675
         Width           =   5550
      End
      Begin VB.TextBox txtServerServerSlotName 
         Height          =   330
         Left            =   1350
         TabIndex        =   11
         Top             =   1080
         Width           =   5550
      End
      Begin VB.TextBox txtServerClientSlotName 
         Height          =   330
         Left            =   1350
         TabIndex        =   10
         Top             =   1485
         Width           =   5550
      End
      Begin VB.Label Label3 
         Caption         =   "Client"
         Height          =   240
         Left            =   90
         TabIndex        =   17
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label Label4 
         Caption         =   "Server"
         Height          =   240
         Left            =   90
         TabIndex        =   16
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label Label5 
         Caption         =   "ServerServer"
         Height          =   240
         Left            =   90
         TabIndex        =   15
         Top             =   1125
         Width           =   1185
      End
      Begin VB.Label Label6 
         Caption         =   "ServerClient"
         Height          =   240
         Left            =   90
         TabIndex        =   14
         Top             =   1530
         Width           =   1185
      End
   End
   Begin VB.TextBox txtError 
      Height          =   330
      Left            =   1260
      TabIndex        =   6
      Top             =   3960
      Width           =   2985
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check Lincense"
      Height          =   285
      Left            =   45
      TabIndex        =   5
      Top             =   2970
      Width           =   1410
   End
   Begin VB.TextBox txtMess 
      Height          =   330
      Left            =   1260
      TabIndex        =   4
      Top             =   3600
      Width           =   2985
   End
   Begin VB.TextBox txtTotUser 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   2970
      Width           =   465
   End
   Begin VB.TextBox txtAction 
      Height          =   330
      Left            =   1260
      TabIndex        =   0
      Top             =   4320
      Width           =   2985
   End
   Begin proLicence.Licence Licence1 
      Left            =   8280
      Top             =   2925
      _ExtentX        =   820
      _ExtentY        =   900
   End
   Begin VB.Label Label10 
      Caption         =   "Max Users Allowed"
      Height          =   195
      Left            =   3105
      TabIndex        =   23
      Top             =   2970
      Width           =   1545
   End
   Begin VB.Label Label9 
      Caption         =   "User List"
      Height          =   285
      Left            =   6165
      TabIndex        =   21
      Top             =   3150
      Width           =   1275
   End
   Begin VB.Label Label8 
      Caption         =   "Error"
      Height          =   240
      Left            =   90
      TabIndex        =   8
      Top             =   4005
      Width           =   915
   End
   Begin VB.Label Label7 
      Caption         =   "Actions"
      Height          =   240
      Left            =   90
      TabIndex        =   7
      Top             =   4365
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Incomming"
      Height          =   285
      Left            =   90
      TabIndex        =   3
      Top             =   3645
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Total User"
      Height          =   240
      Left            =   1575
      TabIndex        =   1
      Top             =   2970
      Width           =   870
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Licence1.ClientSlotname = Me.txtClientSlotname
    Licence1.ServerClientSlotName = Me.txtServerClientSlotName
    Licence1.ServerServerSlotName = Me.txtServerServerSlotName
    Licence1.ServerSlotName = Me.txtServerSlotName
    Licence1.CheckLicence
    Me.txtTotUser = Licence1.TotalUserFound
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then
        frmInfo.Show 1
    End If
    Caption = App.Title
    Me.txtClientSlotname = Licence1.ClientSlotname
    Me.txtServerClientSlotName = Licence1.ServerClientSlotName
    Me.txtServerServerSlotName = Licence1.ServerServerSlotName
    Me.txtServerSlotName = Licence1.ServerSlotName
End Sub

Private Sub Form_Unload(Cancel As Integer)
  '
    ' Close mailslots.
    ' This is very important.  If left open, you will
    ' fail to create them until the OS is rebooted.
    '
Licence1.Quit
End
End Sub

Private Sub Licence1_Actions(Action As String)
'
'this event isplay all action performed by mailslot
'
If LCase(Action) = "closeing" Then
    MsgBox Action
End If
txtAction = Action
txtAction.Refresh
End Sub


Private Sub Licence1_MailslotTimer()
'Is fired on each interval of the timer = 250 ms
End Sub

Private Sub Licence1_MessageIn(IncommingMessage As String)
'
'This event desplays all incomming messages
'
txtMess = IncommingMessage
If InStr(1, UCase(IncommingMessage), "IDENTITY") Then
    lstUsers.AddItem Mid(txtMess, InStr(1, UCase(IncommingMessage), ":"))
End If
End Sub


Private Sub Licence1_OnCountingUser(UserName As String, TotUsers As Integer)
Dim I As Integer
'username is the name of the computer that ar responding
'Totusers are the count of the total users found

lstUsers.Clear
    For I = 1 To TotUsers
        'Retrieve all users from the user collection
        lstUsers.AddItem Licence1.Users(I)
    Next I
    If Licence1.Users.Count >= Me.txtMaxUserLicence Then
        'Licence over pass maxi,=mum allowed
        Licence1.Quit 'Close All mailslots
        frmInfo.info = "Max Licence in Use" & vbCrLf & _
        "Contact your software dealer for more details"
        'Inform the user about more license
        frmInfo.Show 1
        End 'Terminate
    End If
End Sub




Private Sub Licence1_OnError(Errorstring As String, SystError As String, ErrorNumber As Long)
'errorstring : Mailslot Error description
'Systerror   :System error description
'ErrorNumber:System Error Number
'This event is fired each time an error occured

txtError = Errorstring
End Sub

Private Sub options_Click(Index As Integer)
'\\.\mailslot\......\.....\.....'This methos is used for local mailslot adressing
'\\*\mailslot\......\.....\..... used for network
' Ex:
'   \\*\mailslot\Document\Test\Client\Adminprogram.
'   You van make any structure you want.
'   Do not forget "\\*\mailslot\" + Your structure
'   It is better that you adress every local and client and server.
'   Into similar structure.
'   See case below.
        

Select Case Index
    Case 0
        
        'Use deffault settings
        txtClientSlotname = Licence1.ClientSlotname
        txtServerClientSlotName = Licence1.ServerClientSlotName
        txtServerServerSlotName = Licence1.ServerServerSlotName
        txtServerSlotName = Licence1.ServerSlotName
        options(1).Value = False
        options(2).Value = False
    Case 2
        ' Convert to app title   -  Recommended
        txtClientSlotname = "\\.\mailslot\" & App.Title & "\Client"
        txtServerClientSlotName = "\\*\mailslot\" & App.Title & "\Client"
        txtServerServerSlotName = "\\*\mailslot\" & App.Title & "\Server"
        txtServerSlotName = "\\.\mailslot\" & App.Title & "\Server"
        options(1).Value = False
        options(0).Value = False
    Case 1
        options(0).Value = False
        options(2).Value = False
    End Select
End Sub
