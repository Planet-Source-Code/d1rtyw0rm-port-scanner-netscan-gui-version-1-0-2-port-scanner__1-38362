VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function Ping(ByVal IP As String, ByVal Ttl As Long, TimeOut As Long) As Long
Dim hFile As Long, lpWSAdata As WSAdata
Dim hHostent As Hostent, AddrList As Long
Dim Address As Long, rIP As String
Dim OptInfo As IP_OPTION_INFORMATION
Dim EchoReply As IP_ECHO_REPLY
Dim Cmptr As Long

    Call WSAStartup(&H101, lpWSAdata)

    If GetHostByName(IP + String(64 - Len(IP), 0)) <> SOCKET_ERROR Then
        CopyMemory hHostent.h_name, ByVal GetHostByName(IP + String(64 - Len(IP), 0)), Len(hHostent)
        CopyMemory AddrList, ByVal hHostent.h_addr_list, 4
        CopyMemory Address, ByVal AddrList, 4
    End If
    hFile = IcmpCreateFile()
    If hFile = 0 Then
        MsgBox "Unable to Create File Handle", vbCritical, "Network Error"
        Exit Function
    End If
    OptInfo.Ttl = Ttl
    If IcmpSendEcho(hFile, Address, String(32, "A"), 32, OptInfo, EchoReply, Len(EchoReply) + 8, TimeOut) Then
        rIP = CStr(EchoReply.Address(0)) + "." + CStr(EchoReply.Address(1)) + "." + CStr(EchoReply.Address(2)) + "." + CStr(EchoReply.Address(3))
    Else
        rIP = "N/A"
    End If
    If EchoReply.Status = 0 Then
        Ping = EchoReply.RoundTripTime
    Else
        Ping = -666
    End If
    Call IcmpCloseHandle(hFile)
    Call WSACleanup
End Function

Public Function ResolveHostname(ByVal IpAddress As String) As String

Dim hostip_addr As Long
Dim hostent_addr As Long
Dim newAddr As Long
Dim host As Hostent
Dim strTemp As String
Dim strHost As String * 255

    If SocketsInitialize() Then
        newAddr = inet_addr(IpAddress)
        hostent_addr = gethostbyaddr(newAddr, Len(newAddr), AF_INET)

        If hostent_addr = 0 Then
            SocketsCleanup
            Exit Function
        End If

        RtlMoveMemory host, hostent_addr, Len(host)
        RtlMoveMemory ByVal strHost, host.h_name, 255
        strTemp = strHost
        If InStr(strTemp, Chr(0)) <> 0 Then strTemp = Left(strTemp, InStr(strTemp, Chr(0)) - 1)
        strTemp = Trim(strTemp)
        ResolveHostname = strTemp
        SocketsCleanup

    End If
End Function

Private Function SocketsInitialize() As Boolean

    Dim WSAD As WSAdata
    Dim x As Integer
    Dim szLoByte As String
    Dim szHiByte As String
    Dim szBuf As String
    
    x = WSAStartup(WS_VERSION_REQD, WSAD)
    
   'check for valid response
    If x <> 0 Then

        MsgBox "Windows Sockets for 32 bit Windows " & _
               "environments is not successfully responding."
        Exit Function

    End If
    
   'check that the version of sockets is supported
    If lobyte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
       (lobyte(WSAD.wVersion) = WS_VERSION_MAJOR And _
        hibyte(WSAD.wVersion) < WS_VERSION_MINOR) Then
        
        szHiByte = Trim$(Str$(hibyte(WSAD.wVersion)))
        szLoByte = Trim$(Str$(lobyte(WSAD.wVersion)))
        szBuf = "Windows Sockets Version " & szLoByte & "." & szHiByte
        szBuf = szBuf & " is not supported by Windows " & _
                          "Sockets for 32 bit Windows environments."
        MsgBox szBuf, vbExclamation
        Exit Function
        
    End If
    
   'check that there are available sockets
    If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then

        szBuf = "This application requires a minimum of " & _
                 Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        MsgBox szBuf, vbExclamation
        Exit Function

    End If
    
    SocketsInitialize = True
        
End Function

Private Function SocketsCleanup()

   Dim x As Long
   
  'need to use a var (insread of embedding
  'in the If..Then call) becuse the function
  'returns the error code if failed.
   x = WSACleanup()

   If x <> 0 Then
       MsgBox "Windows Sockets error " & Trim$(Str$(x)) & " occurred in Cleanup.", vbExclamation
   End If
    
End Function

Private Function hibyte(ByVal wParam As Long) As Integer
    hibyte = wParam \ &H100 And &HFF&
End Function

Private Function lobyte(ByVal wParam As Long) As Integer
    lobyte = wParam And &HFF&
End Function

