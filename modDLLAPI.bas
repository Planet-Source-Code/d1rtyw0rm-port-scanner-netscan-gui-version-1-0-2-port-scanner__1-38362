Attribute VB_Name = "modDLLAPI"
Const SOCKET_ERROR = 0
Public Type WSAdata
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type
Public Type Hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Public Type IP_OPTION_INFORMATION
    Ttl As Byte
    Tos As Byte
    Flags As Byte
    OptionsSize As Long
    OptionsData As String * 128
End Type
Public Type IP_ECHO_REPLY
    Address(0 To 3) As Byte
    Status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    data As Long
    Options As IP_OPTION_INFORMATION
End Type

Public Const IP_STATUS_BASE = 11000
Public Const IP_SUCCESS = 0
Public Const IP_BUF_TOO_SMALL = (11000 + 1)
Public Const IP_DEST_NET_UNREACHABLE = (11000 + 2)
Public Const IP_DEST_HOST_UNREACHABLE = (11000 + 3)
Public Const IP_DEST_PROT_UNREACHABLE = (11000 + 4)
Public Const IP_DEST_PORT_UNREACHABLE = (11000 + 5)
Public Const IP_NO_RESOURCES = (11000 + 6)
Public Const IP_BAD_OPTION = (11000 + 7)
Public Const IP_HW_ERROR = (11000 + 8)
Public Const IP_PACKET_TOO_BIG = (11000 + 9)
Public Const IP_REQ_TIMED_OUT = (11000 + 10)
Public Const IP_BAD_REQ = (11000 + 11)
Public Const IP_BAD_ROUTE = (11000 + 12)
Public Const IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
Public Const IP_TTL_EXPIRED_REASSEM = (11000 + 14)
Public Const IP_PARAM_PROBLEM = (11000 + 15)
Public Const IP_SOURCE_QUENCH = (11000 + 16)
Public Const IP_OPTION_TOO_BIG = (11000 + 17)
Public Const IP_BAD_DESTINATION = (11000 + 18)
Public Const IP_ADDR_DELETED = (11000 + 19)
Public Const IP_SPEC_MTU_CHANGE = (11000 + 20)
Public Const IP_MTU_CHANGE = (11000 + 21)
Public Const IP_UNLOAD = (11000 + 22)
Public Const IP_ADDR_ADDED = (11000 + 23)
Public Const IP_GENERAL_FAILURE = (11000 + 50)
Public Const MAX_IP_STATUS = 11000 + 50
Public Const IP_PENDING = (11000 + 255)
Public Const PING_TIMEOUT = 200
Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD = 1
'public Const SOCKET_ERROR = -1

Public Const AF_UNSPEC As Integer = 0                    ' unspecified
Public Const AF_UNIX As Integer = 1                      ' local to host (pipes, portals)
Public Const AF_INET As Integer = 2                     ' internetwork: UDP, TCP, etc.
Public Const AF_IMPLINK As Integer = 3                  ' arpanet imp addresses
Public Const AF_PUP As Integer = 4                      ' pup protocols: e.g. BSP
Public Const AF_CHAOS As Integer = 5                    ' mit CHAOS protocols
Public Const AF_IPX As Integer = 6                      ' IPX and SPX
Public Const AF_NS As Integer = AF_IPX                  ' XEROX NS protocols
Public Const AF_ISO As Integer = 7                      ' ISO protocols
Public Const AF_OSI As Integer = AF_ISO                 ' OSI is ISO
Public Const AF_ECMA As Integer = 8                     ' european computer manufacturers
Public Const AF_DATAKIT As Integer = 9                  ' datakit protocols
Public Const AF_CCITT As Integer = 10                    ' CCITT protocols, X.25 etc
Public Const AF_SNA As Integer = 11                      ' IBM SNA
Public Const AF_DECnet As Integer = 12                   ' DECnet
Public Const AF_DLI As Integer = 13                      ' Direct data link interface
Public Const AF_LAT As Integer = 14                      ' LAT
Public Const AF_HYLINK As Integer = 15                  ' NSC Hyperchannel
Public Const AF_APPLETALK As Integer = 16               ' AppleTalk
Public Const AF_NETBIOS As Integer = 17                  ' NetBios-style addresses
Public Const AF_VOICEVIEW As Integer = 18               ' VoiceView
Public Const AF_FIREFOX As Integer = 19                  ' Protocols from Firefox
Public Const AF_UNKNOWN1 As Integer = 20                 ' Somebody is using this!
Public Const AF_BAN As Integer = 21                     ' Banyan
Public Const AF_ATM As Integer = 22                     ' Native ATM Services
Public Const AF_INET6 As Integer = 23                   ' Internetwork Version 6
Public Const AF_CLUSTER As Integer = 24                 ' Microsoft Wolfpack
Public Const AF_12844 As Integer = 25                   ' IEEE 1284.4 WG AF

Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128

Public Type Inet_address
  Byte4 As Byte
  Byte3 As Byte
  Byte2 As Byte
  Byte1 As Byte
End Type
Public IPLong As Inet_address


Public Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Dim ICMPOPT As ICMP_OPTIONS
'Public pIPe As IP_ECHO_REPLY

Public Declare Function GetHostByName Lib "WSOCK32.DLL" Alias "gethostbyname" (ByVal HostName As String) As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired&, lpWSAdata As WSAdata) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Public Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal HANDLE As Long) As Boolean
Public Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean
Public Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal IpAddress$) As Long
Public Declare Function gethostbyaddr Lib "WSOCK32.DLL" (addr As Long, addrLen As Long, addrType As Long) As Long
Public Declare Sub RtlMoveMemory Lib "KERNEL32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)


Public Const ANYSIZE_ARRAY = 1

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Public Type LUID_AND_ATTRIBUTES
    pLuid As LARGE_INTEGER
    Attributes As Long
End Type

Public Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Public Const TOKEN_ADJUST_PRIVILEGES = 32
Public Const TOKEN_QUERY = 8
Public Const SE_PRIVILEGE_ENABLED As Long = 2
Public Const SE_REMOTE_SHUTDOWN_NAME = "SeRemoteShutdownPrivilege"
Public Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"

Public Enum WinPlateformes
    MaskWin32 = &H0
    MaskWin9x = &H1
    MaskWinNT = &H2
End Enum

Public Enum ShutdownActions
    LOGOFF = 0
    ShutDown = 1
    Reboot = 2
    FORCE = 4
End Enum


