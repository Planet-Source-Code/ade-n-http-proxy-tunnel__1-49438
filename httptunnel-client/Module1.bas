Attribute VB_Name = "Module1"
Global Const STATE_IDLE = 0
Global Const STATE_PROXYCONNECT = 1
Global Const STATE_TUNNELCONNECT = 2
Global Const STATE_HOSTCONNECT = 3
Global Const STATE_HOSTCONNECTING = 4
Global Const STATE_OPEN = 5
Global Const STATE_WAIT = 6
Global Const STATE_LOCALESTABLISH = 7
Global Const STATE_TESTING = 8
Global Const STATE_TESTOPEN = 9

Global Const DIODE_PROXYRX = 0
Global Const DIODE_PROXYTX = 1
Global Const DIODE_LOCALRX = 2
Global Const DIODE_LOCALTX = 3
Global Const DIODE_OK = 4
Global Const DIODE_DATA = 5
Global Const DIODE_DP1 = 6
Global Const DIODE_DP2 = 7
Global Const DIODE_ONLINE = 8
Global Const DIODE_TEST = 9
Global Const DIODE_TUNNEL = 10

Global Const DIODE_OFF = &H4000&
Global Const DIODE_ON = &HFF00&


Global Const AHT_CONNECTOK = "ADEHTTPTUNNELOK"
Global Const AHT_HOSTCONNECTOK = "HOSTCONNECTOK"
Global Const AHT_HOSTCONNECTING = "HOSTCONNECTING"
Global Const AHT_OK = "OK"
Global Const AHT_QUIT = "QUIT"
Global Const AHT_BADCMD = "BADCMD"
Global Const AHT_REMOTECLOSE = "REMCLOSE"
Global Const AHT_LINETEST = "LINETEST"


Global Const AHT_SERVERSTRING = "AHTP/1.1"

Global Const AHT_CMD = "0"
Global Const AHT_DATA = "1"
Global Const AHT_DATAPART1 = "2"
Global Const AHT_DATAPART2 = "3"

'Global Const strTestData = "ABCDEF!#¤%&123456"

' Before/After-First/Last created/optimized by aDe
Function BeforeFirst(sIn, sFirst)
    If InStr(1, sIn, sFirst) Then
        BeforeFirst = Left(sIn, InStr(1, sIn, sFirst) - 1)
    Else
        BeforeFirst = ""
    End If
End Function

Function AfterFirst(sIn, sFirst)
    If InStr(1, sIn, sFirst) Then
        AfterFirst = Right(sIn, Len(sIn) - InStr(1, sIn, sFirst) - (Len(sFirst) - 1))
    Else
        AfterFirst = ""
    End If
End Function

Public Function AfterLast(sFrom, sAfterLast)
    If InStr(1, sFrom, sAfterLast) Then
        AfterLast = Right(sFrom, Len(sFrom) - InStrRev(sFrom, sAfterLast) - (Len(sAfterLast) - 1))
    Else
        AfterLast = ""
    End If
End Function

Public Function BeforeLast(sFrom, sBeforeLast)
    If InStr(1, sFrom, sBeforeLast) Then
        BeforeLast = Left(sFrom, InStrRev(sFrom, sBeforeLast) - 1)
    Else
        BeforeLast = ""
    End If
End Function


