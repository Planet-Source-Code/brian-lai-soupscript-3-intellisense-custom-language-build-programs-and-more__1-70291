Attribute VB_Name = "ModEXECTL"
'[ModCTL]
Option Explicit
Public Const Bar0 As Long = 120
Public MyFormNo As Long
Public DebugLevel As Integer
Public AnyKey As Long  'press any key lol

Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nSize As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function IsThemeActive Lib "uxtheme.dll" () As Boolean
Public Declare Function IsAppThemed Lib "uxtheme.dll" () As Boolean
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Public Declare Function GetLogicalDrives Lib "kernel32" () As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Declare Function NetMessageBufferSend Lib "netapi32.dll" (servername As Any, msgname As Byte, fromname As Any, buf As Byte, ByVal buflen As Long) As Long
Public Declare Function NetServerEnum Lib "netapi32.dll" (ByVal servername As String, ByVal level As Long, Buffer As Long, ByVal prefmaxlen As Long, entriesread As Long, totalentries As Long, ByVal servertype As Long, ByVal domain As String, resumehandle As Long) As Long
Public Declare Function NetApiBufferFree Lib "netapi32.dll" (bufptr As Any) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Declare Function lstrcpyW Lib "kernel32" (ByVal lpszDest As String, ByVal lpszSrc As Long) As Long

Public Declare Function DeactivateWindowTheme Lib "uxtheme" Alias "SetWindowTheme" (ByVal hWnd As Long, Optional ByRef pszSubAppName As String = " ", Optional ByRef pszSubIdList As String = " ") As Integer

Public Type SERVER_INFO_100
  sv100_platform_id As Long
  sv100_name As Long
End Type

Public Type SERVER_INFO_101
  dwPlatformId As Long
  lpszServerName As Long
  dwVersionMajor As Long
  dwVersionMinor As Long
  dwType As Long
  lpszComment As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

'netsend consts
Public Const ERROR_ACCESS_DENIED As Long = 5
Public Const ERROR_BAD_NETPATH As Long = 53
Public Const ERROR_INVALID_PARAMETER As Long = 87
Public Const ERROR_NOT_SUPPORTED As Long = 50
Public Const ERROR_INVALID_NAME As Long = 123
Public Const NERR_BASE As Long = 2100
Public Const NERR_SUCCESS As Long = 0
Public Const NERR_NetworkError As Long = (NERR_BASE + 36)
Public Const NERR_NameNotFound As Long = (NERR_BASE + 173)
Public Const NERR_UseNotFound As Long = (NERR_BASE + 150)

Public Const MAX_PREFERRED_LENGTH As Long = -1
Public Const SV_TYPE_ALL                 As Long = &HFFFFFFFF

Public Const DefaultTmpFileName As String = "soupscript.tmp"
Public Const MAX_PATH = 260

Public Function TXTFileOpen(filePath As String) As String
    On Error Resume Next
    Dim FF As Integer
    Dim REAL As String, tmp As String
    FF = FreeFile
    Open filePath For Input As #FF
        Do
            Line Input #FF, tmp
            REAL = REAL & tmp & vbCrLf
        Loop Until EOF(FF)
    Close #FF
    TXTFileOpen = REAL
End Function


Public Sub TXTFileSave(Text As String, filePath As String)
    On Error Resume Next
    Dim F As Integer
    F = FreeFile
    Open filePath For Output As #F
        Print #F, Text
    Close #F
    Exit Sub
End Sub

Public Function MyVer() As String
    On Error Resume Next
    Dim Buffer2 As String
    Dim PreVer As Integer
    PreVer = App.Minor
    If App.Revision >= 1 Then
        PreVer = PreVer + 1
        Buffer2 = Trim$(Str$(PreVer) & " beta")
    Else
        Buffer2 = Trim$(Str$(PreVer))
    End If
    MyVer = "V." & App.Major & "." & Buffer2
End Function

Public Function First(What As String, Optional NumberOfChars As Integer = 15) As String
    On Error Resume Next
    If Len(What) >= NumberOfChars Then
        First = Left$(What, NumberOfChars) & "..."
    Else
        First = What
    End If
End Function

Public Function Last(What As String, Optional NumberOfChars As Integer = 15) As String
    On Error Resume Next
    If Len(What) >= NumberOfChars Then
        Last = "..." & Right$(What, NumberOfChars)
    Else
        Last = What
    End If
End Function

Public Function GetTempDir() As String
   Dim nSize As Long
   Dim tmp As String
   tmp = Space$(MAX_PATH)
   nSize = Len(tmp)
   Call GetTempPath(nSize, tmp)
   GetTempDir = TrimNull(tmp)
End Function

Public Function TrimNull(Item As String)
   Dim Pos As Long
   Pos = InStr(Item, vbNullChar)
    TrimNull = IIf(Pos, Left$(Item, Pos - 1), Item)
End Function

Public Function GetString(Which As String, Optional SectionNo As Long = 0, Optional Delimiter As String = ",") As String
    On Error Resume Next
    Dim Arr() As String
    Arr = Split(Which, Delimiter)
    GetString = Arr(SectionNo)
End Function

Public Function OpenWebpage(strWebpage As String)
    On Error Resume Next
    ShellExecute 0&, vbNullString, strWebpage, vbNullString, vbNullString, vbNormalFocus
End Function

Public Function SDebug(Optional What As String = "", Optional lvl As Integer = 1, Optional AddLine As Boolean = True)
    On Error Resume Next
    With frmMain.TXTImmediate
        If lvl >= DebugLevel Then
            .Text = .Text & What & IIf(AddLine, vbCrLf, "")
        End If
        .SelStart = Len(.Text)
        DoEvents
    End With
End Function

Public Function FindPath(Parent As String, Optional Child As String, Optional Divider As String = "\") As String
    On Error Resume Next
    If Right$(Parent, 1) = Divider Then Parent = Left$(Parent, Len(Parent) - 1)
    If Left$(Child, 1) = Divider Then Child = Mid$(Child, 2)
    FindPath = Parent & Divider & Child
End Function

Public Function ROT13(What As String) As String
    On Error Resume Next
    Dim Bfr1 As String, Bfr2 As String
    Dim TNum As Long
    Dim I As Long
    If Len(What) > 0 Then
        For I = 1 To Len(What) Step 1
            Bfr1 = Mid$(What, I, 1)
            TNum = Asc(Bfr1)
            If TNum >= 65 And TNum <= 90 Then
                If TNum <= 77 Then
                    TNum = TNum + 13
                Else
                    TNum = TNum - 13
                End If
            ElseIf TNum >= 97 And TNum <= 122 Then
                If TNum <= 109 Then
                    TNum = TNum + 13
                Else
                    TNum = TNum - 13
                End If
            Else 'if the text is not in vicinity
                TNum = TNum 'yea i suppose
            End If
            Bfr2 = Bfr2 & Chr$(TNum)
        Next
        ROT13 = Bfr2
    End If
End Function
