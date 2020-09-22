Attribute VB_Name = "modAPI"
'-= API DECLARATION =-'

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
    
Public Declare Function GetPrivateProfileSection Lib "kernel32" _
    Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" _
   Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
   As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileSection Lib "kernel32" _
    Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, _
    ByVal lpString As String, ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName _
    As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lpFileName As String) As Long
