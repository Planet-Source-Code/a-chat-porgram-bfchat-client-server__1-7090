VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "BFChat-Server"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer wait 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4080
      Top             =   2280
   End
   Begin VB.CheckBox log 
      Caption         =   "Log console"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock listen 
      Left            =   4440
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   316
   End
   Begin MSWinsockLib.Winsock win 
      Index           =   0
      Left            =   4800
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame cons 
      Caption         =   "Console"
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7095
      Begin VB.TextBox Console 
         Enabled         =   0   'False
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   6
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.CommandButton say 
      Caption         =   "say"
      Height          =   315
      Left            =   4680
      TabIndex        =   4
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox talk 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Top             =   2760
      Width           =   5055
   End
   Begin VB.TextBox saythat 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   4455
   End
   Begin VB.CommandButton kick 
      Caption         =   "kick him"
      Height          =   255
      Left            =   5280
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
   End
   Begin VB.ListBox peoplehere 
      Height          =   2205
      Left            =   5280
      TabIndex        =   0
      Top             =   2880
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I think all is commented enough ;;;;;-----))))))
'OK: If(when) you got a question feel free to send me an email!
'mail to: visual.basic@gmx.de
' a little tip: Make the Window with th code as wide as possible!!!
' some words to people who are unhappy/happy(whatever) to be german:
'  an die wenigen jungen hoffnungsvollen deutschen Programmierer:
'   1. Du bist nix, ich bin der Chef!
'   2.*Fluch* schon wieder in der Falschen Datei.(Ich schreib neben her einen
'                                                 Brief an meinen Chef!)
'   1. Last euch nicht vom code abschrecken!
'   2. Probiert immer eure Ideen durchzuführen, auch wenns unmöglich erscheint.
'   3. Habt Spaß beim Programmieren(wie ich)!
'   4. Schert euch, wenn ihr programmiert, nicht darum, ob irgendjemand das gut
'      findet, was ihr programmiert.(so wie ich !-))
'Now some words about the program:
'six Chatters are enough! (0, 1, 2, 3, 4, 5)
'there some problems i solved yet
'i think i will solv them and then more then 10 chatter would be possible!
' It works like this:
'
'        connectionrequest
' client------------------>server
'
'           Accepts it.
' client<------------------server
'
'           IP/Name
' client------------------>server
'
' Server closes the socket.
'
'        connectionrequest
' server------------------>client
'
'           Accepts it.
' server<------------------client
'
'         Some shit to chat.
' server<------------------client
'
'  Distributs the shit to all clients.
' server------------------>other clients
'
'Allready asleep? ;-)
'NOW HAVE FUN WITH THE CODE!!!

Private Type chatter  'not realy necessary but cool
 Name As String
 IP As String
End Type
Dim noc As Integer
Dim justconnecting As Integer
Dim chatter(5) As chatter  'think 6 are enough for your memory ;-)
Dim justtext As String
Dim ws(5)                  'for knowing what Winsocks are free(1=used, 0=free,2 = something else)
Dim jIP As String
Dim jname As String
Dim freews As Integer

Private Sub Console_Change()
If log.Value = 1 Then
 Open "c:\bfc_server_console.log" For Output As #1
  Print #1, Console.Text
 Close #1
End If
End Sub

Private Sub Form_Load()
wait.Enabled = False
listen.listen
For i = 1 To 5
Load win(i)
Next i
noc = 0
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If noc = 0 Then Exit Sub
 For i = 0 To 5
   If ws(i) = 1 Then
    If win(i).State = sckConnected Then
    ws(i) = 2  'for sendcomplete()
    win(i).SendData "Server shutdown!"
    DoEvents
    End If
   End If
 Next i
End Sub

Private Sub kick_Click()
 'If MsgBox("Are you sure?", vbYesNo) = vbNo Then Exit Sub  'if you like it, but it halts the server as long as it stays open!!!!
 tokick = peoplehere.List(peoplehere.ListIndex)
 If tokick = "" Then Exit Sub
 If noc = 0 Then Exit Sub
 For i = 0 To 5
  If chatter(i).Name = tokick Then
  If win(i).State = sckConnected Then
   win(i).SendData "You were kicked from the server!"
   peoplehere.RemoveItem (peoplehere.ListIndex)
   ws(i) = 2  'for sendcomplete()
   noc = noc - 1
   GoTo gogogo
  Else
   peoplehere.RemoveItem (peoplehere.ListIndex)
   noc = noc - 1
   GoTo gogogo
  End If
  End If
 Next i
 Console.Text = Console.Text + vbNewLine + "Chatter to be kicked not found!"
gogogo:
End Sub

Private Sub listen_ConnectionRequest(ByVal requestID As Long)
 listen.Close
 listen.Accept requestID
End Sub

Private Sub listen_DataArrival(ByVal bytesTotal As Long)
 listen.GetData dat$
 listen.Close        'done? then close and
 listen.listen       'listen again
 Open "c:\um.tmp" For Output As #1  'not the best, but the easiest way(i was just too lazy to do an better one)
  Print #1, dat$
 Close #1
 Open "c:\um.tmp" For Input As #1
  Input #1, IP$
  Input #1, nam
 Close #1
 If nam = "" Then Exit Sub
 For i = 0 To 5
  If nam = chatter(i).Name Then Exit Sub
 Next i
 If noc = 6 Then GoTo errTOOmanySOCKS 'if you changed the number of chatters, please set this to # of chatters + 1
 For i = 0 To 5
  If ws(i) = 0 Then
  freews = i
  GoTo gogo
  End If
 Next i
 GoTo errTOOmanySOCKS
gogo:
 On Error GoTo errIP
 wait.Enabled = True
 jIP = IP
 jname = nam
 Exit Sub
 
errTOOmanySOCKS:
 Console.Text = "->ERROR: SERVER IS FULL!!!" + vbNewLine + Console.Text
Exit Sub
errIP:
 Console.Text = "The Client " + nam + ", " + IP + " has an bad IP or isn't listening" + vbNewLine + Console.Text
End Sub

Private Sub say_Click()
If saythat.Text = "" Then Exit Sub
  justtext = "@SERVER: " + saythat.Text + vbNewLine + talk.Text  'add it to talk
  talk.Text = "@SERVER: " + saythat.Text + vbNewLine + talk.Text
  For i = 0 To 5
  If ws(i) <> 1 Then GoTo gogogo
   win(i).SendData "@SERVER: " + saythat.Text   'distribut the stuff
   DoEvents
gogogo:
  Next i
  saythat.Text = ""
End Sub



Private Sub saythat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call say_Click
End Sub

Private Sub talk_Change()
talk.Text = justtext  'dont let the user change the content
End Sub

Private Sub wait_Timer()
win(freews).Connect jIP, 317  'ReConnect to client
ws(freews) = 1
chatter(freews).IP = jIP
chatter(freews).Name = jname
peoplehere.AddItem jname
noc = noc + 1
Console.Text = jname + " - " + jIP + " connected" + vbNewLine + Console.Text
wait.Enabled = False
End Sub

Private Sub win_Close(Index As Integer)
ws(Index) = 0
End Sub

Private Sub win_Connect(Index As Integer)
win(Index).SendData vbNewLine + "Welcome on THE ULTIMATE CHATSERVER" + vbNewLine + "BigFchat by MV visual.basic@gmx.de" + vbNewLine
End Sub

Private Sub win_DataArrival(Index As Integer, ByVal bytesTotal As Long)
win(Index).GetData dat$ 'get the stuff and
  If dat$ = "IgogogoNOW!!!" Then
  Call disc(Index)
  Exit Sub
  End If
  justtext = chatter(Index).Name + ": " + dat$ + vbNewLine + talk.Text 'add it to talk
  talk.Text = chatter(Index).Name + ": " + dat$ + vbNewLine + talk.Text
  For i = 0 To 5
  If ws(i) <> 1 Then GoTo gogogo
  On Error GoTo err
   win(i).SendData chatter(Index).Name + ": " + dat$  'distribut the stuff
   DoEvents
   GoTo gogogo
err:
   Console.Text = "Unexcepted error! Chatter:" + chatter(i).Name + " - " + chatter(i).IP + vbNewLine + Console.Text
   Exit For
gogogo:
  Next i
 Exit Sub
End Sub

Private Sub win_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox CStr(Index) + vbNewLine + CStr(Number) + vbNewLine + Description
ws(Index) = 0
win(Index).Close
End Sub

Private Sub win_SendComplete(Index As Integer)
If ws(Index) = 2 Then
 win(Index).Close
 ws(Index) = 0
 Console.Text = chatter(Index).Name + " - " + chatter(Index).IP + " kicked!" + vbNewLine + Console.Text
 chatter(Index).IP = ""
 chatter(Index).Name = ""
End If
End Sub

Private Sub disc(Index)
win(Index).Close
ws(Index) = 0
Console.Text = chatter(Index).Name + " - " + chatter(Index).IP + " disconnected!" + vbNewLine + Console.Text
chatter(Index).IP = ""
chatter(Index).Name = ""
noc = noc - 1
On Error GoTo n
peoplehere.RemoveItem (Index)
n:
End Sub
