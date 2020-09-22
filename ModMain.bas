Attribute VB_Name = "modMain"
Option Explicit

Sub Main()
    On Error Resume Next
    Dim K As String
    
    K = GetMyCode 'Command$
    K = Trim$(K)
    If Len(K) > 0 Then
        'frmMain.Show
        RunCode K
    End If
    End
End Sub

Function RunCode(WhatCode As String)
    On Error Resume Next
    Dim cLine() As String
    Dim TText As String, TA As String, Tb As String
    Dim I As Long, Time1 As Long, Time2 As Long
    
    DebugLevel = 1 'set this up by default
    
    ReDim myPointers(1000) As JustPointers '1000 ought to be enough for everybody - more than enough for 2008 man.
    ReDim myVars(1000) As VarPointers '1000 ought to be enough for everybody - more than enough for 2008 man.
    ReDim myArrays(1000) As ArrayPointers
    
    LoadConsts
    
    Dim TS As New ClsFunctionParser
    'WhatCode is already CodeFriendly
    TS.Code = WhatCode 'CodeFriendly(WhatCode) 'see the CodeFriendly
    TS.RunCode
    Set TS = Nothing
End Function

Function LoadConsts()
    AssignVar "Information", 64, eInteger
    AssignVar "Exclamation", 32, eInteger
    AssignVar "Critical", 16, eInteger
    
    AssignVar "Host", ComputerName, eString
    AssignVar "UserName", UserName, eString
    AssignVar "Command", Command$, eString
    
    'showWindow consts
    AssignVar "Hide", 0, eInteger
    AssignVar "Show", 1, eInteger
    AssignVar "Maximize", 3, eInteger
    AssignVar "Minimize", 2, eInteger
    
    'colors
    AssignVar "White", vbWhite, eLong
    AssignVar "Black", vbBlack, eLong
    AssignVar "Red", vbRed, eLong
    AssignVar "Blue", vbBlue, eLong
    AssignVar "Green", vbGreen, eLong
    
    'var types
    AssignVar "Variant", 0, eInteger
    AssignVar "String", 1, eInteger
    AssignVar "Integer", 2, eInteger
    AssignVar "Long", 3, eInteger
    AssignVar "Double", 4, eInteger
    AssignVar "Date", 5, eInteger
    AssignVar "Boolean", 6, eInteger
    
    AssignVar "pi", 3.141592654, eDouble
End Function

Function GetMyCode() As String
    On Error Resume Next
    Dim BeginPos As Long, I As Long
    Dim varTemp As Variant
    Dim FF As Integer
    Dim byteArr() As Byte
    Dim TC As String
    
    FF = FreeFile
    
    Open FindPath(App.Path, App.EXEName) & ".exe" For Binary As #FF
        Dim FileSize As Long
        Dim FileData As String, FileChunk As String, BFR As String
        FileSize = LOF(FF)
        FileData = Space$(FileSize)
        Get #FF, , FileData
        
        For I = 1 To FileSize Step 1
            If Mid(FileData, I, 6) = "[SOUP]" Then
                I = I + 6
                FileChunk = Mid$(FileData, I) 'Space$(LOF(FF) - I) '
                FileChunk = Trim$(FileChunk) 'remove spaces, if any
                'Get #FF, I, FileChunk
                BFR = FileChunk
                'Exit Function
                Exit For
            End If
        Next I
    Close #FF
    GetMyCode = BFR
End Function
