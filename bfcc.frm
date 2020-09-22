VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "BFChat - Client"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton refresh 
      Height          =   1035
      Left            =   8160
      TabIndex        =   13
      ToolTipText     =   "Refresh IP and Hostname"
      Top             =   2640
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Interval        =   8000
      Left            =   7560
      Top             =   4440
   End
   Begin VB.TextBox nam 
      Height          =   285
      Left            =   6240
      TabIndex        =   11
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox host 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6240
      TabIndex        =   10
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox myip 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6240
      TabIndex        =   8
      Top             =   2640
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   7920
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton disconnect 
      Caption         =   "Disconnect"
      Height          =   615
      Left            =   6360
      TabIndex        =   6
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton connect 
      Caption         =   "Connect"
      Height          =   615
      Left            =   6360
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox Server 
      Height          =   285
      Left            =   6360
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox sayt 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   4560
      Width           =   5175
   End
   Begin VB.CommandButton say 
      Caption         =   "say"
      Height          =   255
      Left            =   5280
      TabIndex        =   1
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox talk 
      Height          =   4455
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   0
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label Status 
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   4920
      Width           =   4695
   End
   Begin VB.Label Label5 
      Caption         =   "Status:"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Chatname:"
      Height          =   255
      Left            =   6240
      TabIndex        =   12
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "MY hostname:"
      Height          =   255
      Left            =   6240
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "MY IP:"
      Height          =   255
      Left            =   6240
      TabIndex        =   7
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Server:"
      Height          =   255
      Left            =   6360
      TabIndex        =   4
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () _
        As Long
        
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal _
        wVersionRequired&, lpWSAData As WinSocketDataType) _
        As Long
        
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () _
        As Long
        
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal _
        HostName$, ByVal HostLen%) As Long
        
Private Declare Function gethostbyname Lib "WSOCK32.DLL" _
        (ByVal HostName$) As Long
        
Private Declare Function gethostbyaddr Lib "WSOCK32.DLL" _
        (ByVal addr$, ByVal laenge%, ByVal typ%) As Long
        
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As _
        Any, ByVal hpvSource&, ByVal cbCopy&)

Const WS_VERSION_REQD = &H101
Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&

Const MIN_SOCKETS_REQD = 1
Const SOCKET_ERROR = -1
Const WSADescription_Len = 256
Const WSASYS_Status_Len = 128


Private Type HostDeType
  hName As Long
  hAliases As Long
  hAddrType As Integer
  hLength As Integer
  hAddrList As Long
End Type

Private Type WinSocketDataType
   wversion As Integer
   wHighVersion As Integer
   szDescription(0 To WSADescription_Len) As Byte
   szSystemStatus(0 To WSASYS_Status_Len) As Byte
   iMaxSockets As Integer
   iMaxUdpDg As Integer
   lpszVendorInfo As Long
End Type
Dim justtext  As String
Dim dat As String
Dim justc As Integer
Dim justd As Integer

Private Sub connect_Click()
ws.Close
On Error GoTo badserver
ws.connect Server.Text, 316
connect.Enabled = False
Server.Enabled = False
Status.Caption = "Connecting... please wait!"
Exit Sub
badserver:
MsgBox "Bad Server! Not found.", vbCritical, "BFC"
connect.Enabled = True
disconnect.Enabled = False
Server.Enabled = True
ws.Close
End Sub

Private Sub disconnect_Click()
ws.SendData "IgogogoNOW!!!"
justd = 1
justc = 0
End Sub

Private Sub Form_Load()
Status.Caption = "Ready"
Timer1.Enabled = False
say.Enabled = False
disconnect.Enabled = False
host.Text = MyHostName
myip.Text = HostByName(MyHostName)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If ws.State = sckConnected Then
Cancel = 1
ws.SendData "IgogogoNOW!!!"
justd = 2
End If
If ws.State = sckConnecting Then
Cancel = 1
MsgBox "Please try again closing the program in a few seconds!" + vbNewLine + "Just connecting to server!"
End If
If ws.State = sckListening Then
Cancel = 1
MsgBox "Please try again closing the program in a few seconds!" + vbNewLine + "Just connecting to server!"
End If
End Sub

Private Sub refresh_Click()
host.Text = MyHostName
myip.Text = HostByName(MyHostName)
End Sub

Private Sub say_Click()
If sayt.Text = "" Then Exit Sub
ws.SendData sayt.Text
sayt.Text = ""
End Sub

Private Sub sayt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call say_Click           'if enter then say
End Sub

Private Sub Server_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call connect_Click       'if enter then connect
End Sub

Private Sub talk_Change()
talk.Text = justtext
End Sub

Private Sub Timer1_Timer()
ws.Close
connect.Enabled = True
disconnect.Enabled = False
Server.Enabled = True
justc = 0
End Sub

Private Sub ws_Connect()
If nam.Text <> "" Then ws.SendData myip.Text + vbNewLine + nam.Text
If nam.Text = "" Then ws.SendData myip.Text + vbNewLine + host.Text  'if no chatname is given, return Computername
Timer1.Enabled = True
justc = 1
End Sub

Private Sub ws_ConnectionRequest(ByVal requestID As Long)
Timer1.Enabled = False
ws.Close
ws.LocalPort = 0
ws.Accept requestID
Status.Caption = "Connected to " + ws.RemoteHostIP
disconnect.Enabled = True
say.Enabled = True
justc = 0
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
ws.GetData dat$
If dat$ = "You were kicked from the server!" Then Call disc
If dat$ = "Server shutdown!" Then Call disc
justtext = dat$ + vbNewLine + talk.Text
talk.Text = dat$ + vbNewLine + talk.Text
End Sub
Private Function MyHostName() As String
  Dim HostName As String * 256
  
    If gethostname(HostName, 256) = SOCKET_ERROR Then
      MsgBox "Windows Sockets error " & Str(WSAGetLastError())
      Exit Function
    Else
      MyHostName = NextChar(Trim$(HostName), Chr$(0))
    End If
End Function

Private Function NextChar(Text$, Char$) As String
  Dim POS%
    POS = InStr(1, Text, Char)
    If POS = 0 Then
      NextChar = Text
      Text = ""
    Else
      NextChar = Left$(Text, POS - 1)
      Text = Mid$(Text, POS + Len(Char))
    End If
End Function

Private Function HostByName(Name$, Optional X% = 0) As String
  Dim MemIp() As Byte
  Dim Y%
  Dim HostDeAddress&, HostIp&
  Dim IpAddress$
  Dim host As HostDeType
  
    HostDeAddress = gethostbyname(Name)
    If HostDeAddress = 0 Then
      HostByName = ""
      Exit Function
    End If
    
    Call RtlMoveMemory(host, HostDeAddress, LenB(host))
    
    For Y = 0 To X
      Call RtlMoveMemory(HostIp, host.hAddrList + 4 * Y, 4)
      If HostIp = 0 Then
        HostByName = ""
        Exit Function
      End If
    Next Y
    
    ReDim MemIp(1 To host.hLength)
    Call RtlMoveMemory(MemIp(1), HostIp, host.hLength)
    
    IpAddress = ""
    
    For Y = 1 To host.hLength
      IpAddress = IpAddress & MemIp(Y) & "."
    Next Y
    
    IpAddress = Left$(IpAddress, Len(IpAddress) - 1)
    HostByName = IpAddress
End Function

Private Sub ws_SendComplete()
If justc = 1 Then
ws.Close
ws.LocalPort = 317
ws.Listen
justc = 0
End If
If justd = 1 Then
ws.Close
connect.Enabled = True
disconnect.Enabled = False
say.Enabled = False
Server.Enabled = True
justd = 0
Status.Caption = "Disconnected and ready"
End If
If justd = 2 Then
ws.Close
End
End If
End Sub

Private Sub disc()
ws.Close
connect.Enabled = True
disconnect.Enabled = False
say.Enabled = False
Server.Enabled = True
End Sub
