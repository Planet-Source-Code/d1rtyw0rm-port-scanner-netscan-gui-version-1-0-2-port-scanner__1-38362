Attribute VB_Name = "modOptions"
Public pngTTL As Long
Public pngTimeOut As Long
Public prtTimeOut As Long
Public prtMaxConn As Long
Public genHost As Boolean
Public genReply As Boolean

Private lpFileName As String

Public Function GetOptions() As Boolean
Dim tmpLen As Long, tmpStr As String * 10
lpFileName = App.Path & "\Conf.ini"
    tmpLen = GetPrivateProfileString("GENERAL", "ResolveHost", "True", tmpStr, Len(tmpStr), lpFileName)
        If InStr(tmpStr, "Vrai") Then
            genHost = True
        Else
            genHost = False
        End If
        
    tmpLen = GetPrivateProfileString("GENERAL", "PingReply", "True", tmpStr, Len(tmpStr), lpFileName)
        If InStr(tmpStr, "Vrai") Then
            genReply = True
        Else
            genReply = False
        End If
              
    tmpLen = GetPrivateProfileString("PING", "TTL", "130", tmpStr, Len(tmpStr), lpFileName)
        pngTTL = CLng(tmpStr)
        
    tmpLen = GetPrivateProfileString("PING", "TimeOut", "300", tmpStr, Len(tmpStr), lpFileName)
        pngTimeOut = CLng(tmpStr)
        
    tmpLen = GetPrivateProfileString("PORT", "MaxConnection", "5", tmpStr, Len(tmpStr), lpFileName)
        prtMaxConn = CLng(tmpStr)
        
    tmpLen = GetPrivateProfileString("PORT", "TimeOut", "1300", tmpStr, Len(tmpStr), lpFileName)
        prtTimeOut = CLng(tmpStr)

End Function

Public Function SaveOptions(ByVal ResolveHost As String, ByVal PingReplyOnly As String, ByVal PingTTL As String, ByVal PingTimeOut As String, ByVal PortMaxConnection As String, ByVal PortTimeOut As String)
lpFileName = App.Path & "\Conf.ini"

    WritePrivateProfileString "GENERAL", "ResolveHost", ResolveHost, lpFileName
    WritePrivateProfileString "GENERAL", "PingReply", PingReplyOnly, lpFileName
    WritePrivateProfileString "PING", "TTL", PingTTL, lpFileName
    WritePrivateProfileString "PING", "TimeOut", PingTimeOut, lpFileName
    WritePrivateProfileString "PORT", "MaxConnection", PortMaxConnection, lpFileName
    WritePrivateProfileString "PORT", "TimeOut", PortTimeOut, lpFileName

End Function
