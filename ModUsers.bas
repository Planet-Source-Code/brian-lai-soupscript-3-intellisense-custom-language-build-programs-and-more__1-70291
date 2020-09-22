Attribute VB_Name = "ModUsers"
Option Explicit

Private Declare Function NetUserEnum Lib "netapi32" (servername As Byte, ByVal level As Long, ByVal filter As Long, buff As Long, ByVal buffsize As Long, entriesread As Long, totalentries As Long, resumehandle As Long) As Long
Private Declare Function GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Function GetComputersName() As String
    On Error Resume Next
  'returns the name of the computer
   Dim tmp As String
   tmp = Space$(16)
   If GetComputerName(tmp, Len(tmp)) <> 0 Then
      GetComputersName = TrimNull(tmp)
   End If
End Function

Public Function UserName() As String
    On Error Resume Next
    Dim lpBuffer As String
    Dim J
    lpBuffer = Space$(255)
    GetUserName lpBuffer, Len(lpBuffer)
        J = InStr(lpBuffer, Chr$(0))
    If J > 0 Then UserName = Left$(lpBuffer, J - 1)
End Function

Public Function ComputerName() As String
    On Error Resume Next
    Dim tmp As String
    tmp = Space$(16)
    If GetComputerName(tmp, Len(tmp)) <> 0 Then
        ComputerName = TrimNull(tmp)
    End If
End Function

Public Function LoadServers(Optional ArrayPrefix As String = "NetUsers", Optional sDomain As String = vbNullString) As Long
    On Error Resume Next
    'source: VBNet
    
    Dim bufptr As Long, dwEntriesread As Long, dwTotalentries As Long, _
    dwResumehandle As Long, success As Long, nStructSize As Long, CNT As Long
    
    Dim se100 As SERVER_INFO_100
    Dim StrOutput As String
    
    nStructSize = LenB(se100)
    SDebug "ArrayPrefix = " & ArrayPrefix & ". Please wait...", 0
    success = NetServerEnum(StrOutput, 100, bufptr, MAX_PREFERRED_LENGTH, dwEntriesread, dwTotalentries, SV_TYPE_ALL, 0, dwResumehandle)
    'if all goes well
    If success = NERR_SUCCESS Then
        For CNT = 0 To dwEntriesread - 1
            CopyMemory se100, ByVal bufptr + (nStructSize * CNT), nStructSize
            StrOutput = GetPointerToByteStringW(se100.sv100_name)
            AssignVar CStr(ArrayPrefix & CNT + 1), StrOutput, eString
            SDebug "Assigned variable " & ArrayPrefix & CNT + 1 & " to " & StrOutput, 0
            DoEvents
        Next
    Else
        SDebug "Error: success = " & success, 2
    End If
    'clean up regardless of success
    Call NetApiBufferFree(bufptr)
    LoadServers = dwEntriesread
End Function

Public Function GetPointerToByteStringW(ByVal dwData As Long) As String
    On Error Resume Next
    Dim tmp() As Byte
    Dim tmplen As Long
    If dwData <> 0 Then
    tmplen = lstrlenW(dwData) * 2
        If tmplen <> 0 Then
           ReDim tmp(0 To (tmplen - 1)) As Byte
           CopyMemory tmp(0), ByVal dwData, tmplen
           GetPointerToByteStringW = tmp
        End If
    End If
End Function
