VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ade http tunnel - server"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8355
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrClientTimeout 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   5640
      Top             =   120
   End
   Begin VB.Timer tmrClosecheck 
      Interval        =   500
      Left            =   5160
      Top             =   120
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   3960
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdToggle 
      Caption         =   "Start/Stop"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtPort 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   "5500"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtConsole 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   8055
   End
   Begin MSWinsockLib.Winsock SocketR 
      Left            =   4560
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "local port:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------
' 2003-10-25
' Welcome to the source code of 'ade http tunnel - server'
' As usual there's few to no comments - good luck =)
'
' Feel free to use the code as you please as long as you don't try
' to steal my credit etc etc.
'
' aDe
' ade@ade.se
'----------------------------------------------------------------------

Dim bActive As Boolean, bRemConnected As Boolean, sDataBuffer(0 To 99) As String, iDataType(0 To 99) As String * 1
Dim sInBuffer As String, cliInBuffer As String

Private Sub cmdToggle_Click()
ToggleState
End Sub
Sub ToggleState()
If bActive Then
  Socket.Close
  SocketR.Close
  bActive = False
  CleanBuffers
  cOut "Server shutdown OK"
  tmrClientTimeout.Enabled = False
  bRemConnected = False
Else
  Socket.Close
  Socket.LocalPort = txtPort.Text
  Socket.Listen
  bActive = True
  cOut "Server started on port " & Socket.LocalPort
  SaveSettings
End If
End Sub

Sub CleanBuffers()
For i = 0 To UBound(sDataBuffer)
  sDataBuffer(i) = ""
Next
End Sub

Sub cOut(sOut As String)
txtConsole.Text = txtConsole.Text & "[" & Time & "] " & sOut & vbCrLf
txtConsole.SelStart = Len(txtConsole)
If Len(txtConsole) > 65000 Then txtConsole = ""

End Sub

Private Sub Form_Load()
LoadSettings
End Sub

Private Sub Form_Unload(Cancel As Integer)
If bActive Then ToggleState
SaveSettings

End Sub

Private Sub Socket_Close()
If bActive Then
cOut "Proxy disconnected. Resetting"
ToggleState
ToggleState
End If
End Sub

Private Sub Socket_ConnectionRequest(ByVal requestID As Long)
Socket.Close
Socket.Accept requestID
cOut "Client " & Socket.RemoteHostIP & " connected"
tmrClientTimeout.Enabled = True
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
Dim sIn As String, sAct As String, sData As String, sDataLen As String, sDataType As String
If Socket.State <> sckConnected Then
  Socket.Close
  Exit Sub
End If
Socket.GetData sIn
tmrClientTimeout.Enabled = False
tmrClientTimeout.Enabled = True

If Left(sIn, 5) = "POST " Then
  If AfterFirst(sIn, vbCrLf & vbCrLf) = "" Then
    sInBuffer = sIn
    'cOut "proxy indata buffered..."
    Exit Sub
  End If
End If

If Len(sInBuffer) Then
  sIn = sInBuffer & sIn
  sInBuffer = ""
End If

If Left(sIn, 3) = "GET" Then
  sAct = BeforeFirst(AfterFirst(sIn, "GET "), " HTTP/")
ElseIf Left(sIn, 5) = "POST " Then
  sAct = BeforeFirst(AfterFirst(sIn, "POST "), " HTTP/")
  If Left(LCase(sAct), 7) = "http://" Then sAct = AfterLast(sAct, "/")
  sDataLen = BeforeFirst(AfterFirst(sIn, "Content-Length: "), vbCrLf)
  sData = Right(sIn, sDataLen - 2)
  sDataType = AfterLast(BeforeFirst(sIn, "=" & sData), vbCrLf)
End If


If Left(sAct, 1) = "/" Then sAct = Right(sAct, Len(sAct) - 1)

Select Case sAct
  Case "connect"
    MakeReply AHT_CONNECTOK, AHT_CMD
  Case "open"
    cOut "Client requested connection"
    RemoteConnect sData
  Case "refresh"
    If Not bRemConnected Then
      cOut "Client sent invalid /refresh request; dropping"
      MakeReply AHT_BADCMD, AHT_CMD
      DoEvents
      Call Socket_Close
      Exit Sub
    End If
    refreshClient
  Case "send"
    If Not bRemConnected Then
      cOut "Client sent invalid /send request; dropping"
      MakeReply AHT_BADCMD, AHT_CMD
      DoEvents
      Call Socket_Close
      Exit Sub
    End If
    If sDataType = AHT_DATAPART1 Then
      cliInBuffer = sData
      refreshClient
    ElseIf sDataType = AHT_DATAPART2 Then
      cliInBuffer = cliInBuffer & sData
      rSend sData
      refreshClient
    ElseIf sDataType = AHT_DATA Then
      rSend sData
      refreshClient
    End If
  Case AHT_QUIT
    cOut "Client logging off."
    MakeReply AHT_OK, AHT_CMD
    DoEvents
    
    Call Socket_Close
  Case AHT_LINETEST
    MakeReply AHT_LINETEST & sData, AHT_CMD
End Select
End Sub

Sub refreshClient()
    If Len(sDataBuffer(0)) Then
      MakeReply sDataBuffer(0), iDataType(0), IIf(Len(sDataBuffer(1)), "refresh=1", "")
      For i = 0 To 98
        sDataBuffer(i) = sDataBuffer(i + 1)
        iDataType(i) = iDataType(i + 1)
      Next
      sDataBuffer(99) = ""
    Else
      MakeReply AHT_OK, AHT_CMD
    End If
End Sub

Sub rSend(sData As String)
If SocketR.State <> sckConnected Then
  cOut "rSend: No connection"
  Exit Sub
End If
SocketR.SendData sData
End Sub

Sub RemoteConnect(sServer As String)
Dim sAddr As String, sPort As String

sAddr = BeforeFirst(sServer, ":")
sPort = AfterFirst(sServer, ":")

If Len(sAddr) = 0 Then
  cOut """/open"" request invalid."
  Exit Sub
End If

SocketR.Close
SocketR.Connect sAddr, sPort
cOut "Connecting to " & sAddr & ":" & sPort & "..."
MakeReply AHT_HOSTCONNECTING, AHT_CMD
End Sub
Sub MakeReply(sData As String, sType As String, Optional sCookies As String)
Dim sOut As String
sOut = sOut & "HTTP/1.1 200 OK" & vbCrLf
sOut = sOut & "Server: " & AHT_SERVERSTRING & vbCrLf
'sOut = sOut & "Date: " & Now & vbCrLf
sOut = sOut & "Connection: Keep-Alive" & vbCrLf
sOut = sOut & "Content-Length: " & (Len(sData) + 1) & vbCrLf
sOut = sOut & "Content-Type: text/html" & vbCrLf
If Len(sCookies) Then sOut = sOut & "Set-Cookie: " & sCookies & vbCrLf
sOut = sOut & "Cache-control: No-Cache" & vbCrLf
sOut = sOut & vbCrLf
sOut = sOut & sType & sData
'cOut ">>type " & sType
sockSend sOut
End Sub
Sub sockSend(sData As String)
If Socket.State <> sckConnected Then cOut "sockSend: No connection": Exit Sub
Socket.SendData sData
End Sub

Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
cOut "Socket error: " & Description

End Sub

Private Sub SocketR_Close()
If bActive Then
  cOut "Remote host closed connection."
  saveToDataBuf AHT_REMOTECLOSE, AHT_CMD
End If
End Sub

Private Sub SocketR_Connect()
saveToDataBuf AHT_HOSTCONNECTOK, AHT_CMD
bRemConnected = True
cOut "Connected to host."
End Sub

Sub saveToDataBuf(sSave As String, Optional iType As Integer = AHT_DATA)
For i = 0 To 99
  If Len(sDataBuffer(i)) = 0 Then
    sDataBuffer(i) = sSave
    iDataType(i) = iType
    'cOut "saveToDataBuf(" & iType & ")"
    Exit For
  End If
Next
End Sub

Private Sub SocketR_DataArrival(ByVal bytesTotal As Long)
Dim sIn As String, sDP1 As String, sDP2 As String, sL As Integer
If SocketR.State <> sckConnected Then
  SocketR.Close
  Exit Sub
End If
SocketR.GetData sIn
If Len(sIn) Then
  Do While Len(sIn)
    If Len(sIn) > 1000 Then
      sDP1 = Left(sIn, 1000)
      sIn = Right(sIn, Len(sIn) - 1000)
      
      sL = Len(sIn) - 1000
      If sL < 0 Then
        sDP2 = sIn
        sIn = ""
      Else
        sDP2 = Left(sIn, 1000)
        sIn = Right(sIn, Len(sIn) - 1000)
      End If
      saveToDataBuf sDP1, AHT_DATAPART1
      saveToDataBuf sDP2, AHT_DATAPART2
    Else
      saveToDataBuf sIn, AHT_DATA
      sIn = ""
    End If
  Loop
End If
End Sub

Private Sub tmrClientTimeout_Timer()
If bActive Then cOut "Client timeout": ToggleState: ToggleState
End Sub

Private Sub tmrClosecheck_Timer()
If SocketR.State = 8 Then
  SocketR.Close
  Exit Sub
End If
If Socket.State = 8 Then
  Socket.Close
End If
End Sub


Sub SaveSettings()
SaveSetting "aDe", "aht-server", "listen", txtPort.Text

End Sub

Sub LoadSettings()
txtPort.Text = GetSetting("aDe", "aht-server", "listen", txtPort.Text)

End Sub
