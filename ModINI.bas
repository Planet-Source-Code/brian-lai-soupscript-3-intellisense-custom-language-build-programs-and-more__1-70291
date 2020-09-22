Attribute VB_Name = "ModINI"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal sSectionName As String, ByVal sReturnedString As String, ByVal lSize As Long, ByVal sFileName As String) As Long
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Function ReadINI(Section As String, KeyName As String, FileName As String) As String
On Error Resume Next
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName$, "", sRet, Len(sRet), FileName))
End Function

Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
On Error Resume Next
    Dim r
    r = writeprivateprofilestring(sSection, sKeyName, sNewString, sFileName)
End Function

Function GetSet(Key As String, Optional Default As String, Optional NodeName As String = "Program") As String
    On Error Resume Next
    Dim Buffer As String, Buffer2 As String
            
    Buffer = ReadINI(NodeName, Key, SettingsFile) 'read the entry from my username section
    If Len(Buffer) > 0 Then 'if there's an entry in my user name
        GetSet = Buffer 'then so be it
        GoTo ExitFunction
    End If
    
    If Len(Default) > 0 Then 'if nothing else is present but I have a defined preset...
        GetSet = Default 'then so be it
        GoTo ExitFunction
    End If

Exit Function

ExitFunction:
End Function

Function SaveSet(Key As String, Value As String, Optional NodeName As String = "Program") As String
On Error Resume Next
    WriteINI NodeName, Key, Value, SettingsFile
    SaveSet = Key
End Function

Public Function SettingsFile() As String
    On Error Resume Next
    SettingsFile = FindPath(App.Path, App.ProductName & ".ini")
End Function



