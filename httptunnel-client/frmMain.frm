VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ade http tunnel - client"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSockRef 
      Interval        =   2000
      Left            =   1200
      Top             =   3000
   End
   Begin MSWinsockLib.Winsock SockL 
      Left            =   240
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock SockR 
      Left            =   720
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox frameSettings 
      BackColor       =   &H00785028&
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1875
      ScaleWidth      =   6555
      TabIndex        =   1
      Top             =   120
      Width           =   6615
      Begin VB.CommandButton cmdToggleTest 
         Caption         =   "Test tunnel"
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
         Left            =   3650
         TabIndex        =   15
         Top             =   1425
         Width           =   1080
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1830
         Left            =   0
         Picture         =   "frmMain.frx":0E42
         ScaleHeight     =   1830
         ScaleWidth      =   855
         TabIndex        =   14
         Top             =   0
         Width           =   855
         Begin VB.Shape shapeDiode 
            BorderColor     =   &H0000FF00&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00004000&
            FillStyle       =   0  'Solid
            Height          =   90
            Index           =   10
            Left            =   495
            Top             =   1610
            Width           =   135
         End
         Begin VB.Shape shapeDiode 
            BorderColor     =   &H0000FF00&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00004000&
            FillStyle       =   0  'Solid
            Height          =   90
            Index           =   9
            Left            =   495
            Top             =   1340
            Width           =   135
         End
         Begin VB.Shape shapeDiode 
            BorderColor     =   &H0000FF00&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00004000&
            FillStyle       =   0  'Solid
            Height          =   90
            Index           =   8
            Left            =   495
            Top             =   1755
            Width           =   135
         End
         Begin VB.Shape shapeDiode 
            BorderColor     =   &H0000FF00&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00004000&
            FillStyle       =   0  'Solid
            Height          =   90
            Index           =   7
            Left            =   495
            Top             =   1185
            Width           =   135
         End
         Begin VB.Shape shapeDiode 
            BorderColor     =   &H0000FF00&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00004000&
            FillStyle       =   0  'Solid
            Height          =   90
            Index           =   6
            Left            =   495
            Top             =   1035
            Width           =   135
         End
         Begin VB.Shape shapeDiode 
            BorderColor     =   &H0000FF00&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00004000&
            FillStyle       =   0  'Solid
            Height          =   90
            Index           =   5
            Left            =   495
            Top             =   885
            Width           =   135
         End
         Begin VB.Shape shapeDiode 
            BorderColor     =   &H0000FF00&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00004000&
            FillStyle       =   0  'Solid
            Height          =   90
            Index           =   4
            Left            =   495
            Top             =   735
            Width           =   135
         End
         Begin VB.Shape shapeDiode 
            BorderColor     =   &H0000FF00&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00004000&
            FillStyle       =   0  'Solid
            Height          =   90
            Index           =   3
            Left            =   735
            Top             =   330
            Width           =   135
         End
         Begin VB.Shape shapeDiode 
            BorderColor     =   &H0000FF00&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00004000&
            FillStyle       =   0  'Solid
            Height          =   90
            Index           =   2
            Left            =   495
            Top             =   330
            Width           =   135
         End
         Begin VB.Shape shapeDiode 
            BorderColor     =   &H0000FF00&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00004000&
            FillStyle       =   0  'Solid
            Height          =   90
            Index           =   1
            Left            =   735
            Top             =   180
            Width           =   135
         End
         Begin VB.Shape shapeDiode 
            BorderColor     =   &H0000FF00&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00004000&
            FillStyle       =   0  'Solid
            Height          =   90
            Index           =   0
            Left            =   495
            Top             =   180
            Width           =   135
         End
      End
      Begin VB.CommandButton cmdSetRefresh 
         Caption         =   "Set"
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
         Left            =   6015
         TabIndex        =   13
         Top             =   135
         Width           =   450
      End
      Begin VB.CommandButton cmdToggle 
         Caption         =   "Start/stop service"
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
         Left            =   4785
         TabIndex        =   12
         Top             =   1425
         Width           =   1680
      End
      Begin VB.TextBox txtRefresh 
         Appearance      =   0  'Flat
         BackColor       =   &H00804000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5325
         TabIndex        =   10
         Text            =   "1000"
         Top             =   135
         Width           =   660
      End
      Begin VB.Timer tmrDiode 
         Enabled         =   0   'False
         Index           =   0
         Interval        =   500
         Left            =   3840
         Top             =   120
      End
      Begin VB.TextBox txtProxy 
         Appearance      =   0  'Flat
         BackColor       =   &H00804000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2610
         TabIndex        =   5
         Top             =   450
         Width           =   3855
      End
      Begin VB.TextBox txtHost 
         Appearance      =   0  'Flat
         BackColor       =   &H00804000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2610
         TabIndex        =   4
         Top             =   1080
         Width           =   3855
      End
      Begin VB.TextBox txtLocalPort 
         Appearance      =   0  'Flat
         BackColor       =   &H00804000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2610
         TabIndex        =   3
         Text            =   "1200"
         Top             =   135
         Width           =   975
      End
      Begin VB.TextBox txtTunnel 
         Appearance      =   0  'Flat
         BackColor       =   &H00804000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2610
         TabIndex        =   2
         Text            =   "http://"
         Top             =   765
         Width           =   3855
      End
      Begin VB.Label lblGui 
         BackColor       =   &H00785028&
         Caption         =   "Refresh time (ms):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   3885
         TabIndex        =   11
         Top             =   150
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000F&
         X1              =   1020
         X2              =   1020
         Y1              =   0
         Y2              =   1875
      End
      Begin VB.Label lblGui 
         BackColor       =   &H00785028&
         Caption         =   "Proxy server:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   1170
         TabIndex        =   9
         Top             =   450
         Width           =   1095
      End
      Begin VB.Label lblGui 
         BackColor       =   &H00785028&
         Caption         =   "Remote host:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1170
         TabIndex        =   8
         Top             =   1095
         Width           =   1095
      End
      Begin VB.Label lblGui 
         BackColor       =   &H00785028&
         Caption         =   "Local listening port:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   1170
         TabIndex        =   7
         Top             =   150
         Width           =   1455
      End
      Begin VB.Label lblGui 
         BackColor       =   &H00785028&
         Caption         =   "Tunnel address:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   1170
         TabIndex        =   6
         Top             =   765
         Width           =   1170
      End
   End
   Begin VB.TextBox txtConsole 
      BackColor       =   &H0070A3A8&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2160
      Width           =   6615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------
' 2003-10-25
' Welcome to the source code of 'ade http tunnel - client'
' As usual there's few to no comments - good luck =)
'
' Feel free to use the code as you please as long as you don't try
' to steal my credit etc etc.
'
' aDe
' ade@ade.se
'----------------------------------------------------------------------

Dim bActive As Boolean, iState As Integer, sDataBuffer(0 To 99) As String, iDataType(0 To 99) As String * 1
Dim pInBuffer As String, sLastReq As String, timeLastReq As String, retryCount As Integer
Dim bLoggingOff As Boolean, bTunnelConnected As Boolean, bTestMode As Boolean
Dim objTimer As New clsTimer
Dim strTestData As String, iTestCount As Integer



Private Sub cmdSetRefresh_Click()
tmrSockRef.Interval = txtRefresh.Text
cOut "Refresh rate set: " & tmrSockRef.Interval & " ms"
End Sub

Private Sub cmdToggle_Click()
If Not bActive Then
  StartUp
Else
  ShutDown
End If
End Sub

Sub cleanDataBuffer()
For i = 0 To UBound(sDataBuffer)
  sDataBuffer(i) = ""
  iDataType(i) = ""
Next
End Sub

Sub ShutDown()

If bTestMode Then
  cmdToggle.Enabled = True
  bTestMode = False
  cmdToggleTest.Caption = "Test tunnel"
Else
  SockL.Close
End If

If bTunnelConnected = True Then
  cOut "Logging off server..."
  ProxyRequest AHT_QUIT
End If
DoEvents
SockR.Close
iState = STATE_IDLE
tmrSockRef.Enabled = False
bActive = False
cOut "Server shutdown OK"
shapeDiode(DIODE_ONLINE).FillColor = DIODE_OFF
shapeDiode(DIODE_TUNNEL).FillColor = DIODE_OFF
cleanDataBuffer
retryCount = 0
bLoggingOff = False
bTunnelConnected = False
End Sub

Sub StartUp()
Dim rAddr As String, rPort As String
On Error GoTo errhndl
GoTo runsub
errhndl:
ErrHandle Error$
Exit Sub
runsub:

If bTestMode Then
  cmdToggle.Enabled = False
  cmdToggleTest.Caption = "Stop test"
Else
  SockL.LocalPort = txtLocalPort.Text
  SockL.Listen
  cOut "Waiting for local connection."
End If

If InStr(1, txtProxy.Text, ":") = 0 Then
  cOut "No proxy port; assuming 3128"
  txtProxy.Text = txtProxy.Text & ":3128"
End If
rAddr = BeforeFirst(txtProxy.Text, ":")
rPort = AfterFirst(txtProxy.Text, ":")
SockR.RemoteHost = rAddr
SockR.RemotePort = rPort
cOut "Server has been started."

tmrSockRef.Enabled = True
bActive = True

If bTestMode Then ProxyConnect
End Sub

Sub ErrHandle(sErrdesc As String)
cOut "Error: " & sErrdesc
End Sub

Sub cOut(sOut As String)
txtConsole.Text = txtConsole.Text & "[" & Time & "] " & sOut & vbCrLf
txtConsole.SelStart = Len(txtConsole)
If Len(txtConsole) > 65000 Then txtConsole = ""

End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub cmdToggleTest_Click()
If bTestMode Then
  ShutDown
Else
  If Len(strTestData) = 0 Then
    For i = 0 To 255
      strTestData = strTestData & Chr(i)
    Next
  End If
  iTestCount = 0
  bTestMode = True
  StartUp
End If
End Sub

Private Sub Form_Load()
For i = 1 To shapeDiode.UBound
  Load tmrDiode(i)
Next

LoadSettings
Caption = "aDe http tunnel - client v" & App.Major & "." & App.Minor & App.Revision
Call cmdSetRefresh_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSettings
SockL.Close
SockR.Close

End Sub

Sub DiodeBlink(Index As Integer)
  shapeDiode(Index).FillColor = DIODE_ON
  tmrDiode(Index).Enabled = False
  tmrDiode(Index).Enabled = True
  DoEvents
End Sub

Sub SaveSettings()
  SaveSetting "aDe", "ahp", "localport", txtLocalPort.Text
  SaveSetting "aDe", "ahp", "host", txtHost.Text
  SaveSetting "aDe", "ahp", "proxy", txtProxy.Text
  SaveSetting "aDe", "ahp", "tunnel", txtTunnel.Text
  SaveSetting "aDe", "ahp", "refresh", txtRefresh.Text
End Sub

Sub LoadSettings()
 txtLocalPort.Text = GetSetting("aDe", "ahp", "localport", txtLocalPort.Text)
 txtHost.Text = GetSetting("aDe", "ahp", "host", txtHost.Text)
 txtProxy.Text = GetSetting("aDe", "ahp", "proxy", txtProxy.Text)
 txtTunnel.Text = GetSetting("aDe", "ahp", "tunnel", txtTunnel.Text)
 txtRefresh.Text = GetSetting("aDe", "ahp", "refresh", txtRefresh.Text)
End Sub

Private Sub SockL_Close()
If bActive Then
  cOut "Local connection lost"
  ResetServer
End If
End Sub

Private Sub SockL_ConnectionRequest(ByVal requestID As Long)
cOut "Local client is trying to connect..."
SockL.Close
iState = STATE_PROXYCONNECT
ProxyConnect


Do Until (iState = STATE_LOCALESTABLISH) Or (iState = STATE_IDLE)
  DoEvents
Loop

If iState = STATE_LOCALESTABLISH Then
  SockL.Close
  SockL.Accept requestID
  cOut "Local connection established"
  shapeDiode(DIODE_ONLINE).FillColor = DIODE_ON
  iState = STATE_OPEN
Else
  cOut "Local connection cancelled."
End If
End Sub

Private Sub SockL_DataArrival(ByVal bytesTotal As Long)
Dim sIn As String, sDP1 As String, sDP2 As String, sL As Integer
DiodeBlink DIODE_LOCALRX
SockL.GetData sIn
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

If iState = STATE_OPEN Then
  Call tmrSockRef_Timer
End If
End Sub

Sub saveToDataBuf(sSave As String, Optional iType As Integer = AHT_DATA)
If bLoggingOff Then Exit Sub
For i = 0 To 99
  If Len(sDataBuffer(i)) = 0 Then
    sDataBuffer(i) = sSave
    iDataType(i) = iType
    Exit For
  End If
Next
End Sub

Private Sub SockL_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
ErrHandle Description
End Sub

Sub ProxyConnect()
cOut "Connecting to proxy..."
iState = STATE_PROXYCONNECT
SockR.Connect

End Sub

Private Sub SockR_Close()
If bActive Then
  cOut "Proxy connection lost"
  ResetServer
End If
End Sub

Sub ResetServer()
If bTestMode Then
  ShutDown
Else
  ShutDown
  StartUp
End If
End Sub

Private Sub SockR_Connect()
Dim sOut As String
cOut "Connected to proxy."
cOut "Connecting to tunnel..."
iState = STATE_TUNNELCONNECT
ProxyRequest "connect"

End Sub
Sub ProxyRequest(Optional sAdd As String)
Dim sA As String, sOut As String
If Len(sAdd) Then sA = sAdd
If LCase(Left(txtTunnel.Text, 7)) <> "http://" Then txtTunnel.Text = "http://" & txtTunnel.Text
If Right(txtTunnel.Text, 1) <> "/" Then txtTunnel.Text = txtTunnel.Text & "/"

sOut = "GET " & txtTunnel.Text & sA & " HTTP/1.0" & vbCrLf
sOut = sOut & "Accept: */*" & vbCrLf
sOut = sOut & "Accept-Language: en" & vbCrLf
sOut = sOut & "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)" & vbCrLf
sOut = sOut & "Host: " & BeforeFirst(AfterFirst(LCase(txtTunnel), "http://"), "/") & vbCrLf
sOut = sOut & "Proxy-Connection: Keep-Alive" & vbCrLf
sOut = sOut & vbCrLf
rSend sOut
End Sub

Sub ProxyPost(sData As String, sAdd As String, Optional nDataType As String = AHT_DATA)
Dim sOut As String, sendLen As Integer
sendLen = Len(sData) + 2


If Len(sAdd) Then sA = sAdd
If LCase(Left(txtTunnel.Text, 7)) <> "http://" Then txtTunnel.Text = "http://" & txtTunnel.Text & vbCrLf
If Right(txtTunnel.Text, 1) <> "/" Then txtTunnel.Text = txtTunnel.Text & "/" & vbCrLf

sOut = "POST " & txtTunnel.Text & sA & " HTTP/1.0" & vbCrLf
sOut = sOut & "Accept: */*" & vbCrLf
'sOut = sOut & "Accept-Language: en" & vbCrLf
'sOut = sOut & "Referer: " & txtTunnel.Text & "connect" & vbCrLf
sOut = sOut & "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)" & vbCrLf
sOut = sOut & "Host: " & BeforeFirst(AfterFirst(LCase(txtTunnel), "http://"), "/") & vbCrLf
sOut = sOut & "Proxy-Connection: Keep-Alive" & vbCrLf
sOut = sOut & "Content-Type: application/x-www-form-urlencoded" & vbCrLf
sOut = sOut & "Content-Length: " & sendLen & vbCrLf
sOut = sOut & "Pragma: no-cache" & vbCrLf
sOut = sOut & vbCrLf
sOut = sOut & nDataType & "=" & sData

Clipboard.Clear
Clipboard.SetText sOut
rSend sOut
End Sub
Private Sub SockR_DataArrival(ByVal bytesTotal As Long)
Dim sIn As String, rOK As Boolean, sData As String, cLen As String, sType As String
Dim sHeadData As String
Dim sCookies As String, srvStr As String

DiodeBlink DIODE_PROXYRX

If SockR.State <> sckConnected Then Exit Sub

SockR.GetData sIn
'cOut sIn

retryCount = 0

If UCase(Left(sIn, 15)) = "HTTP/1.1 200 OK" Then
  rOK = True
ElseIf UCase(Left(sIn, 15)) = "HTTP/1.0 200 OK" Then
  rOK = True
End If

sHeadData = BeforeFirst(sIn, vbCrLf & vbCrLf)

cLen = BeforeFirst(AfterFirst(sHeadData, "Content-Length: "), vbCrLf)
If InStr(1, sHeadData, "Set-Cookie: ") Then
  sCookies = BeforeFirst(AfterFirst(sHeadData, "Set-Cookie: "), vbCrLf)
End If

srvStr = BeforeFirst(AfterFirst(sHeadData, "Server: "), vbCrLf)
If srvStr <> AHT_SERVERSTRING Then
  cOut "Warning: serverstring mismatch. (" & srvStr & ")"
End If

If Not rOK Then Exit Sub
sData = Right(sIn, Val(cLen))
sType = Left(sData, 1)
sData = Right(sData, Len(sData) - 1)
'cOut "data: " & sIn
If sType = AHT_CMD Then
  If bTestMode Then
    If Left(sData, Len(AHT_LINETEST)) = AHT_LINETEST Then
      objTimer.StopTimer
      DiodeBlink DIODE_TEST
      sData = Right(sData, Len(sData) - Len(AHT_LINETEST))
      If sData <> strTestData Then
        cOut "<< (" & Replace(Format(objTimer.Elapsed, "0.0"), ",", ".") & " ms) " & Len(sData) & " bytes, data ERROR"
      Else
        cOut "<< (" & Replace(Format(objTimer.Elapsed, "0.0"), ",", ".") & " ms) " & Len(sData) & " bytes, data OK!"
      End If
      iTestCount = iTestCount + 1
      iState = STATE_TESTOPEN
      If iTestCount = 5 Then ShutDown
      Exit Sub
    End If
  End If
  If sData = AHT_BADCMD Then
    cOut "Server returned BADCMD, disconnecting"
    Call SockR_Close
  End If
  If sData = AHT_REMOTECLOSE Then
    cOut "Remote host connection closed."
    ShutDown
    Exit Sub
  End If
  Select Case iState
    Case STATE_TUNNELCONNECT
      If sData = AHT_CONNECTOK Then
        cOut "Connected to tunnel."
        shapeDiode(DIODE_TUNNEL).FillColor = DIODE_ON
        bTunnelConnected = True
        If bTestMode Then
          cOut "Testing with " & Len(strTestData) & " bytes of data:"
          SendTestData
        Else
          iState = STATE_HOSTCONNECT
          ProxyPost txtHost.Text, "open"
          cOut "Sending connect to host request..."
        End If
      End If
    Case STATE_HOSTCONNECT
      If sData = AHT_HOSTCONNECTING Then
        cOut "Server is connecting to host..."
        iState = STATE_HOSTCONNECTING
      End If
    Case STATE_HOSTCONNECTING
        If sData = AHT_HOSTCONNECTOK Then
          cOut "Host connection successfull."
          iState = STATE_LOCALESTABLISH
        End If
    Case STATE_WAIT
        If sData = AHT_OK Then
          DiodeBlink DIODE_OK
          If bLoggingOff Then
            If sDataBuffer(0) = "" Then cOut "Successfully logged off.": ShutDown
          End If
        End If
        iState = STATE_OPEN
  End Select
ElseIf sType = AHT_DATA Then
  'cOut "Data"
  DiodeBlink DIODE_DATA
  lSend sData
  If sCookies = "refresh=1" Then
    iState = STATE_OPEN
    Call tmrSockRef_Timer
  Else
    iState = STATE_OPEN
  End If
ElseIf sType = AHT_DATAPART1 Then
  'cOut "Datapart1"
  DiodeBlink DIODE_DP1
  pInBuffer = sData
  iState = STATE_OPEN
  Call tmrSockRef_Timer
ElseIf sType = AHT_DATAPART2 Then
  'cOut "Datapart2"
  DiodeBlink DIODE_DP2
  pInBuffer = pInBuffer & sData
  lSend pInBuffer
  pInBuffer = ""
  iState = STATE_OPEN
End If
End Sub

Sub SendTestData()
  ProxyPost strTestData, AHT_LINETEST
  objTimer.ResetTimer
  iState = STATE_TESTING
End Sub
Sub rSend(sData As String)
If SockR.State = sckConnected Then
  SockR.SendData sData
  DiodeBlink DIODE_PROXYTX
  'Clipboard.Clear
  'Clipboard.SetText sData
  sLastReq = sData
  timeLastReq = Now
Else
  cOut "Error rSend: No connection"
End If
End Sub
Sub lSend(ByVal sData As String)
If SockL.State = sckConnected Then
  DiodeBlink DIODE_LOCALTX
  SockL.SendData sData
Else
  cOut "Error lSend: No connection"
End If
End Sub


Private Sub tmrDiode_Timer(Index As Integer)
  shapeDiode(Index).FillColor = DIODE_OFF
  tmrDiode(Index).Enabled = False
  DoEvents
End Sub

Private Sub tmrSockRef_Timer()
If SockR.State = 8 Then SockR.Close

If bTestMode Then
  If iState = STATE_TESTOPEN Then
    SendTestData
  End If
  Exit Sub
End If

If (iState = STATE_OPEN) Or (iState = STATE_HOSTCONNECTING) Then
  If iState = STATE_OPEN Then iState = STATE_WAIT
  
  If Len(sDataBuffer(0)) Then
    ProxyPost sDataBuffer(0), "send", iDataType(0)
    For i = 0 To 98
      sDataBuffer(i) = sDataBuffer(i + 1)
      iDataType(i) = iDataType(i + 1)
    Next
    sDataBuffer(99) = ""
    iDataType(99) = ""
  Else
    ProxyRequest "refresh"
  End If
ElseIf iState = STATE_WAIT Then
  If DateDiff("s", timeLastReq, Now) >= 5 Then
    If retryCount >= 5 Then
      cOut "Retried 5 times, aborting!"
      cmdToggle_Click
      cmdToggle_Click
      Exit Sub
    End If
    cOut "No response recieved, retrying"
    retryCount = retryCount + 1
    rSend sLastReq
  End If
End If
End Sub
