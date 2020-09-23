VERSION 5.00
Begin VB.UserControl Licence 
   Appearance      =   0  'Flat
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Licence.ctx":0000
   ScaleHeight     =   480
   ScaleWidth      =   450
   ToolboxBitmap   =   "Licence.ctx":030A
   Begin VB.Timer Timer_Client 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   495
      Top             =   630
   End
End
Attribute VB_Name = "Licence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer$, nSize As Long) As Long
Option Explicit

Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

Private Declare Function CreateMailslot Lib "kernel32" Alias "CreateMailslotA" (ByVal lpName As String, ByVal nMaxMessageSize As Long, ByVal lReadTimeout As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function CreateMailslotNoSecurity Lib "kernel32" Alias "CreateMailslotA" (ByVal lpName As String, ByVal nMaxMessageSize As Long, ByVal lReadTimeout As Long, ByVal Zero As Long) As Long
Private Declare Function GetMailslotInfo Lib "kernel32" (ByVal hMailslot As Long, lpMaxMessageSize As Long, lpNextSize As Long, lpMessageCount As Long, lpReadTimeout As Long) As Long
Private Declare Function SetMailslotInfo Lib "kernel32" (ByVal hMailslot As Long, ByVal lReadTimeout As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function ReadFileSimple Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal Zero As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function WriteFileSimple Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal Zero As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFileNoSecurity Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal Zero As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Const MAILSLOT_NO_MESSAGE = (-1)
Const MAILSLOT_WAIT_FOREVER = (-1)
Const GENERIC_WRITE = &H40000000
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const OPEN_EXISTING = 3

' ServerSlot is the "handle" of the server mailslot.
Event OnError(Errorstring As String, SystError As String, ErrorNumber As Long)
Event Actions(Action As String)
Event MailslotTimer()
Event OnCountingUser(UserName As String, TotUsers As Integer)
Event MessageIn(IncommingMessage As String)
'Default Property Values:
Const m_def_CountTimeDelay = 3
Const m_def_TotalUserFound = 0
Const m_def_ServerSlot = 0
Const m_def_ClientSlot = 0
Const m_def_ServerSlotName = "\\.\mailslot\YourAppTitle\Server"
Const m_def_ClientSlotname = "\\.\mailslot\YourAppTitle\Client"
Const m_def_ServerClientSlotname = "\\*\mailslot\YourAppTitle\Client"
Const m_def_ServerServerSlotname = "\\*\mailslot\YourAppTitle\Server"
'Property Variables:
Dim m_CountTimeDelay As Integer
Dim m_TotalUserFound As Integer
Dim m_ServerSlot As Long
Dim m_ClientSlot As Long
Dim m_ServerSlotName As String

' ServerSlotName is the pathname of the mailslot file.
' It must begin with "\\" and be followed by the machine
' name or "." for local, and must be followed by "\mailslot\".
' After that, you can make up any valid directory tree and
' filename.  The client must know this name in order to
' start a conversation with the server.

Dim m_ClientSlotname As String
' ClientSlotName is the name of the client's mailslot.
' ClientSlot is the "handle" of an open client.
' It is used to talk to the client.
Dim m_ServerClientSlotname As String
Dim m_ServerServerSlotname As Variant
Public Users As New Collection



Function MailSlotRead(SlotHandle As Long) As String
    '
    ' Read a specified mailslot.
    ' Return the text of the next message (if any).
    '
    Dim Result As Long ' result of system calls
    Dim MessageNew As String ' text of new message
    Dim MessageCount As Long ' count of messages waiting in slot
    Dim MessageLength As Long ' len of next messgae in queue
    Dim ReadTimeout As Long ' how long to wait for message
    Dim BytesRead As Long ' actual count of bytes read from slot
    ' Check if any messages are waiting in mailslot
    Result = GetMailslotInfo(SlotHandle, 0, MessageLength, MessageCount, 0)
    If Result = 0 Then
        RaiseEvent OnError("Bad return from GetMailslotInfo", Err.Description, Err.Number)
        Exit Function
    End If
    If MessageLength = MAILSLOT_NO_MESSAGE Then
        MailSlotRead = ""
        Exit Function
    End If
    ' Retrieve next message from mailslot
    MessageNew = String$(MessageLength + 1, " ")
    Result = ReadFileSimple(SlotHandle, MessageNew, MessageLength, BytesRead, 0)
    If Result = 0 Then
        RaiseEvent OnError("Failed to read message", Err.Description, Err.Number)
        MailSlotRead = ""
        Exit Function
    End If
    If BytesRead <> MessageLength Then
        RaiseEvent OnError("Did not read correct message length", Err.Description, Err.Number)
        MailSlotRead = ""
        Exit Function
    End If
    ' Return message
    MailSlotRead = MessageNew
    RaiseEvent MessageIn(MessageNew)
End Function

Sub MailSlotWrite(SlotHandle As Long, Text As String)
    '----------------------------------------------------------------------------------------------------
    ' Write message to server's mailslot.
    '----------------------------------------------------------------------------------------------------
        
    Dim Result As Boolean ' result of system calls
    Dim TextLen As Long ' length of message to send
    Dim BytesWritten As Long ' actual count of bytes sent

    '----------------------------------------------------------------------------------------------------
    ' Send message over mailslot
    '----------------------------------------------------------------------------------------------------
    
    TextLen = Len(Text) + 1
    Result = WriteFileSimple(SlotHandle, Text, TextLen, BytesWritten, 0)
    If Not Result Then
        RaiseEvent OnError("failed to write message", Err.Description, Err.Number)
        Exit Sub
    End If
    If BytesWritten <> TextLen Then
        RaiseEvent OnError("wrote wrong number of bytes", Err.Description, Err.Number)
        Exit Sub
    End If
End Sub



Function ComputerName() As String
    Dim sTmp1 As String
    sTmp1 = Space(512)
    GetComputerName sTmp1, Len(sTmp1)
    ComputerName = Trim(sTmp1)
End Function

Private Sub Timer_Client_Timer()
'----------------------------------------------------------------------------------------------------
'active when a client, this look up for messages
'----------------------------------------------------------------------------------------------------
        
    Dim sText As String
    sText = MailSlotRead(ClientSlot)
    
    If Len(sText) > 0 Then
        ' display new message
        
        If InStr(1, UCase(sText), "IDENTIFY") > 0 Then
            RaiseEvent Actions("Identifying")
            Call MailSlotWrite(ServerSlot, "IDENTITY:" & ComputerName)
            Beep
        End If
        RaiseEvent MessageIn(sText)
        RaiseEvent MailslotTimer
    End If
End Sub

Private Sub UserControl_Terminate()
RaiseEvent Actions("Closing")
Quit
End Sub


Public Sub CheckLicence(Optional MaxUserLicence As Integer = 1)
On Error GoTo Out
    'Start Application
    'step 1: Try to be a Mailslot server
    ' OK  : Request how many users are using your application accros the network
    ' FAIL: Become a client
Dim Result As Boolean
    '----------------------------------------------------------------------------------------------------
    ' Reset All slots before starting
    '----------------------------------------------------------------------------------------------------
    If ServerSlot > 0 Then Result = CloseHandle(ServerSlot)
    If ClientSlot > 0 Then Result = CloseHandle(ClientSlot)
    
    If Create_Server Then
        '----------------------------
        'Start as MailslotServer
        '----------------------------
        RaiseEvent Actions("Server Started")
        
        Get_Server_ClientSlot
        '----------------------------------------------------------------------------------------------------
        'now request users working with the application across the network
        '----------------------------------------------------------------------------------------------------
        If ClientSlot > 0 Then
            TotalUserFound = User_Count
            If TotalUserFound >= MaxUserLicence Then
                RaiseEvent Actions("All Licence In Use")
            End If
            RaiseEvent Actions("Closing Server")
            Result = CloseHandle(ClientSlot)
            Result = CloseHandle(ServerSlot)
            
            ClientSlot = 0
            ServerSlot = 0
'----------------------------------------------------------------------------------------------------
'       Now become a Client to receive messege
'----------------------------------------------------------------------------------------------------
            Get_Client_ClientSlot
            Get_ServerSlot
            Timer_Client.Enabled = True
            Timer_Client.Interval = 250
        Else
            RaiseEvent OnError("Unknown Mailslot error", Err.Description, Err.Number)
            '----------------------------------------------------------------------------------------------------
            'there is a problem
            '----------------------------------------------------------------------------------------------------
        End If
    Else
        
        '----------------------------------------------------------------------------------------------------
        'Client
        '----------------------------------------------------------------------------------------------------
        RaiseEvent Actions("Client Started")
        Get_Client_ClientSlot
        Timer_Client.Interval = 250
    End If
Exit Sub
Out:
End Sub

Private Function User_Count() As Integer
Dim iAantal As Integer
Dim Text As String
Dim iTeller As Integer
Dim dtijd As Date
Dim Status As String
Dim I As Integer
Dim uCount As Integer
Dim Found As Boolean
    Call MailSlotWrite(ClientSlot, "IDENTIFY:")
    '=====================
    'Wait for incomming messages
    '=====================
    RaiseEvent Actions("Counting Users")
    Status = "Counting Users"
    dtijd = Now()
    I = CountTimeDelay
    '=============================
    'Wait some times to receive the response
    '=============================
    While DateDiff("s", dtijd, Now()) < I
        If Status <> "Counting Users (" & Format((DateDiff("s", dtijd, Now()) / I) * 100, "0") & ")% Completed" Then
            RaiseEvent Actions(Status)
        End If
    Status = "Counting Users (" & Format((DateDiff("s", dtijd, Now()) / I) * 100, "0") & ")% Completed"
    Wend
    RaiseEvent Actions("READY")
        
    '=====================
    ' Check for inbound messages
    '=====================
    iTeller = 0
    Text = MailSlotRead(ServerSlot)
    While Len(Text) > 0
        RaiseEvent MessageIn(Text)
        If InStr(1, Text, "IDENTITY") Then
        '==========================
        'Got a response of other PC running
        'Your program across the network
        '==========================
            iTeller = iTeller + 1
            TotalUserFound = iTeller
            '=======================
            'Fired the OnCountingUser event
            '=======================
            For uCount = 1 To Users.Count - 1
            '
            'Is the user already added?
            '
                Found = (Users.Item(uCount) = Mid(Text, InStr(1, UCase(Text), ":") + 1))
            Next uCount
            If Not Found Then
            '========================
            'Add the user to the collection
            '========================
                Users.Add Mid(Text, InStr(1, UCase(Text), ":") + 1)
                RaiseEvent OnCountingUser(Mid(Text, InStr(1, UCase(Text), ":") + 1), Users.Count)
            End If
        End If
       '
       'Read the next message in the mailslot queue
       '
       Text = MailSlotRead(ServerSlot)
       
    Wend
    User_Count = Users.Count
End Function


Private Function Create_Server() As Boolean
'----------------------------------------------------------------------------------------------------
'Retrieve Serverslot op
'----------------------------------------------------------------------------------------------------
    If ServerSlot = 0 Then
        ServerSlot = CreateMailslotNoSecurity(ServerSlotName, 0, 0, 0)
        If ServerSlot = -1 Then
        '----------------------------------------------------------------------------------------------------
        'Server already exsists!
        '----------------------------------------------------------------------------------------------------
            Create_Server = False
            Exit Function
        Else
            Create_Server = True
        End If
    ElseIf ServerSlot = -1 Then
'----------------------------------------------------------------------------------------------------
'Server already exsists
'----------------------------------------------------------------------------------------------------
                Create_Server = False
        Exit Function
    Else
'----------------------------------------------------------------------------------------------------
'There is a serverslot already, so nothing to do
'----------------------------------------------------------------------------------------------------
                Create_Server = True
    End If
End Function

Private Sub Get_Server_ClientSlot()
'----------------------------------------------------------------------------------------------------
'Now create a server Slot where All clients can identify them selfs
'----------------------------------------------------------------------------------------------------
    If ClientSlot = 0 Then
        ClientSlot = CreateFileNoSecurity(ServerClientSlotname, _
            GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_EXISTING, _
            FILE_ATTRIBUTE_NORMAL, 0)
        If ClientSlot = 0 Then RaiseEvent Actions("Server Started")
    End If
End Sub

Private Sub Get_ServerSlot()
'----------------------------------------------------------------------------------------------------
'Create A serverslot to receive message
'----------------------------------------------------------------------------------------------------
        
    If ServerSlot = 0 Then
        ServerSlot = CreateFileNoSecurity(ServerServerSlotname, GENERIC_WRITE, _
            FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    End If
End Sub

Private Sub Get_Client_ClientSlot()
'----------------------------------------------------------------------------------------------------
' Create A Clientslot to receive message
'----------------------------------------------------------------------------------------------------
                    
    If ClientSlot = 0 Then
        ClientSlot = CreateMailslotNoSecurity(ClientSlotname, 0, 0, 0)
        RaiseEvent Actions("Client Started")
    End If
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,\\.\mailslot\YourAppTitle\Server
Public Property Get ServerSlotName() As String
    ServerSlotName = m_ServerSlotName
End Property

Public Property Let ServerSlotName(ByVal New_ServerSlotName As String)
    m_ServerSlotName = New_ServerSlotName
    PropertyChanged "ServerSlotName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,\\.\mailslot\YourAppTitle\Client"
Public Property Get ClientSlotname() As String
    ClientSlotname = m_ClientSlotname
End Property

Public Property Let ClientSlotname(ByVal New_ClientSlotname As String)
    m_ClientSlotname = New_ClientSlotname
    PropertyChanged "ClientSlotname"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,\\*\mailslot\YourAppTitle\Client"
Public Property Get ServerClientSlotname() As String
    ServerClientSlotname = m_ServerClientSlotname
End Property

Public Property Let ServerClientSlotname(ByVal New_ServerClientSlotname As String)
    m_ServerClientSlotname = New_ServerClientSlotname
    PropertyChanged "ServerClientSlotname"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,\*\mailslot\YourAppTitle\Server
Public Property Get ServerServerSlotname() As Variant
    ServerServerSlotname = m_ServerServerSlotname
End Property

Public Property Let ServerServerSlotname(ByVal New_ServerServerSlotname As Variant)
    m_ServerServerSlotname = New_ServerServerSlotname
    PropertyChanged "ServerServerSlotname"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ServerSlotName = m_def_ServerSlotName
    m_ClientSlotname = m_def_ClientSlotname
    m_ServerClientSlotname = m_def_ServerClientSlotname
    m_ServerServerSlotname = m_def_ServerServerSlotname
    m_ClientSlot = m_def_ClientSlot
    m_ServerSlot = m_def_ServerSlot
    m_TotalUserFound = m_def_TotalUserFound
    m_CountTimeDelay = m_def_CountTimeDelay
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_ServerSlotName = PropBag.ReadProperty("ServerSlotName", m_def_ServerSlotName)
    m_ClientSlotname = PropBag.ReadProperty("ClientSlotname", m_def_ClientSlotname)
    m_ServerClientSlotname = PropBag.ReadProperty("ServerClientSlotname", m_def_ServerClientSlotname)
    m_ServerServerSlotname = PropBag.ReadProperty("ServerServerSlotname", m_def_ServerServerSlotname)
    m_ClientSlot = PropBag.ReadProperty("ClientSlot", m_def_ClientSlot)
    m_ServerSlot = PropBag.ReadProperty("ServerSlot", m_def_ServerSlot)
    m_TotalUserFound = PropBag.ReadProperty("TotalUserFound", m_def_TotalUserFound)
    m_CountTimeDelay = PropBag.ReadProperty("CountTimeDelay", m_def_CountTimeDelay)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("ServerSlotName", m_ServerSlotName, m_def_ServerSlotName)
    Call PropBag.WriteProperty("ClientSlotname", m_ClientSlotname, m_def_ClientSlotname)
    Call PropBag.WriteProperty("ServerClientSlotname", m_ServerClientSlotname, m_def_ServerClientSlotname)
    Call PropBag.WriteProperty("ServerServerSlotname", m_ServerServerSlotname, m_def_ServerServerSlotname)
    Call PropBag.WriteProperty("ClientSlot", m_ClientSlot, m_def_ClientSlot)
    Call PropBag.WriteProperty("ServerSlot", m_ServerSlot, m_def_ServerSlot)
    Call PropBag.WriteProperty("TotalUserFound", m_TotalUserFound, m_def_TotalUserFound)
    Call PropBag.WriteProperty("CountTimeDelay", m_CountTimeDelay, m_def_CountTimeDelay)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,2,0
Public Property Get ClientSlot() As Long
Attribute ClientSlot.VB_MemberFlags = "400"
    ClientSlot = m_ClientSlot
End Property

Public Property Let ClientSlot(ByVal New_ClientSlot As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    m_ClientSlot = New_ClientSlot
    PropertyChanged "ClientSlot"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,2,0
Public Property Get ServerSlot() As Long
Attribute ServerSlot.VB_MemberFlags = "400"
    ServerSlot = m_ServerSlot
End Property

Public Property Let ServerSlot(ByVal New_ServerSlot As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    m_ServerSlot = New_ServerSlot
    PropertyChanged "ServerSlot"
End Property

Public Sub Quit()
  '
    ' Close mailslots.
    ' This is very important.  If left open, you will
    ' fail to create them until the OS is rebooted.
    '
Static Once As Boolean
Dim Result As Boolean
If Once Then Exit Sub
    Result = CloseHandle(ServerSlot)
    Result = CloseHandle(ClientSlot)
RaiseEvent Actions("Closing")
Once = True
UserControl_Terminate
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,2,0
Public Property Get TotalUserFound() As Integer
Attribute TotalUserFound.VB_MemberFlags = "400"
    TotalUserFound = m_TotalUserFound
End Property

Public Property Let TotalUserFound(ByVal New_TotalUserFound As Integer)
    If Ambient.UserMode = False Then Err.Raise 387
    m_TotalUserFound = New_TotalUserFound
    PropertyChanged "TotalUserFound"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,3
Public Property Get CountTimeDelay() As Integer
    CountTimeDelay = m_CountTimeDelay
End Property

Public Property Let CountTimeDelay(ByVal New_CountTimeDelay As Integer)
    m_CountTimeDelay = New_CountTimeDelay
    PropertyChanged "CountTimeDelay"
End Property

