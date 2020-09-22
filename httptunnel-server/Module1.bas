Attribute VB_Name = "Module1"
Global Const AHT_CONNECTOK = "ADEHTTPTUNNELOK"
Global Const AHT_OK = "OK"
Global Const AHT_HOSTCONNECTOK = "HOSTCONNECTOK"
Global Const AHT_HOSTCONNECTING = "HOSTCONNECTING"
Global Const AHT_QUIT = "QUIT"
Global Const AHT_BADCMD = "BADCMD"
Global Const AHT_REMOTECLOSE = "REMCLOSE"
Global Const AHT_LINETEST = "LINETEST"

Global Const AHT_CMD = "0"
Global Const AHT_DATA = "1"
Global Const AHT_DATAPART1 = "2"
Global Const AHT_DATAPART2 = "3"


Global Const AHT_SERVERSTRING = "AHTP/1.1"

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


