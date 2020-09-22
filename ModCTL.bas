Attribute VB_Name = "ModCTL"
'[ModCTL]
Option Explicit
Public Const Bar0 As Long = 120
Public MyFormNo As Long
Public DebugLevel As Integer

Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Sub FatalAppExit Lib "kernel32" Alias "FatalAppExitA" (ByVal uAction As Long, ByVal lpMessageText As String)
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nSize As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Public Declare Function IsThemeActive Lib "uxtheme.dll" () As Boolean
Public Declare Function IsAppThemed Lib "uxtheme.dll" () As Boolean
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Declare Function DeleteUrlCacheEntry Lib "Wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
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
Public Declare Function NetApiBufferFree Lib "netapi32.dll" (BufPtr As Any) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Declare Function lstrcpyW Lib "kernel32" (ByVal lpszDest As String, ByVal lpszSrc As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Type SERVER_INFO_101
  dwPlatformId As Long
  lpszServerName As Long
  dwVersionMajor As Long
  dwVersionMinor As Long
  dwType As Long
  lpszComment As Long
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

Public Const StdFN As String = "SoupBase.dll"
Public Const DefaultTmpFileName As String = "soupscript.tmp"
Public Const MAX_PATH = 260

Public Validation As String 'DRM file path


Function AF() As Form 'returns the active form. if theres no form open, then it will create a new document to prevent errors
    On Error Resume Next
    If Forms.Count > 1 Then
        Set AF = frmMain.ActiveForm
    Else
        Set AF = New frmScript
        AF.Show
    End If
End Function

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

Function SStatus(Optional What As String = "Ready", Optional WarningLevel As Integer = 2)
    On Error Resume Next
    Dim K As Long
    Select Case WarningLevel
        Case 0
            K = vbBlack
        Case 1
            K = RGB(128, 128, 0)
        Case Is >= 2
            K = vbRed
    End Select
    With frmMain.lblStatus
        .Caption = What
        .FontBold = (WarningLevel >= 2)
        .ForeColor = K
    End With
End Function

Function SyncTabs(Optional SyncTabTags As Boolean = True)
    On Error Resume Next
    Dim K As String
    frmMain.Tb.RemoveAllTabs
    Dim Frm As Form
    For Each Frm In Forms
        With frmMain.Tb
            If Frm.Tag <> "CLOSED" And TypeOf Frm Is frmScript Then
                .AddTab IIf(Frm.Caption = "*", "Unsaved*", Last(Frm.Caption))
                .TabTooltip(.TabUBound) = Last(Frm.Caption)
                .TabTag(.TabUBound) = Frm.Tag
                If SyncTabTags Then .ActiveTab = Val(AF.Tag)
            End If
        End With
    Next
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

'Public Function SDebug(What As String, Optional lvl As Integer = 1)
'    On Error Resume Next
'    If lvl < DebugLevel Then Exit Function
''    If lvl >= DebugLevel And AF.pDP.Height <= AF.lblDrag.Height Then AF.pDP.Height = AF.ScaleHeight / 2
'    If lvl >= 2 And AF.pDP.Height <= AF.lblDrag.Height Then AF.pDP.Height = AF.ScaleHeight / 2
'
'    AF.TXTImmediate.Text = AF.TXTImmediate.Text & What & vbCrLf
'    AF.TXTImmediate.SelStart = Len(AF.TXTImmediate.Text)
'End Function

Public Function GetString(Which As String, Optional SectionNo As Long = 0, Optional Delimiter As String = ",") As String
    On Error Resume Next
    Dim Arr() As String
    Arr = Split(Which, Delimiter)
    GetString = Arr(SectionNo)
End Function

Public Sub SkinFormEx(Which As Form)
    On Error Resume Next
    Dim A As Control
    Dim B As Integer, E As Integer, F As Integer, I As Integer
    Dim C As String
    Dim D As Boolean

    B = IIf(IsAppThemed And IsThemeActive, 0, 1)
    If B <> 0 Then
        For Each A In Which
            If Len(A.Name) = 0 Then Exit For
            If B = 1 Then
                If TypeOf A Is CommandButton Then BTFlat A
                CtlFlat A
            End If
        Next
    End If
    If B = 1 Then FormFlat Which
End Sub

Public Sub BTFlat(bt As CommandButton)
    On Error Resume Next
        If GetWindowLong&(bt.hWnd, -16) And &H8000& Then Exit Sub
        SetWindowLong bt.hWnd, -16, GetWindowLong&(bt.hWnd, -16) Or &H8000&
        bt.Refresh
End Sub

Public Sub CtlFlat(CL As Control)
    On Error Resume Next
        CL.Appearance = 0   'flat
        'CL.BackColor = &H8000000F 'looks more natural  'for cham buttons, and they change backcolor to the same as the container
        CL.ColorScheme = 2 'for cham buttons only
        'CL.BackOver = &H8000000F 'looks more natural 'for cham buttons only
End Sub

Public Sub FormFlat(Which As Form)
    On Error Resume Next
    Which.Appearance = 0
    Which.BackColor = &H8000000F 'looks more natural
End Sub

Public Function OpenWebpage(strWebpage As String)
    On Error Resume Next
    ShellExecute 0&, vbNullString, strWebpage, vbNullString, vbNullString, vbNormalFocus
End Function

Public Function GetHint(cmd As String, Optional DoNotShowAnything As Boolean) As String
    On Error Resume Next
    GetHint = GetSet(cmd, IIf(DoNotShowAnything, "", "No help available"), "CodeLib")
End Function

Function PutDLL()
    On Error Resume Next
    'If IsIDE Then Exit Function 'please don't mess with myself when i'm editing.
    Dim I() As Byte
    Dim K As String
    Dim FF As Integer
    FF = FreeFile
    K = FindPath(App.Path, StdFN)
    
    Select Case Val(GetSet("MakeDLL", "0"))
        Case 0 'create if not present
            If Len(Dir(K)) = 0 Then GoTo MakeThisDLL
        Case 1 'create always
            GoTo MakeThisDLL
    End Select
    
    Exit Function
MakeThisDLL:
    I = LoadResData(101, "CUSTOM")
    Open K For Binary Access Write As #FF
        Put #FF, , I
    Close #FF
End Function

Public Function KillEXE(Optional WhatName As String)
    On Error Resume Next
    If Len(WhatName) = 0 Then WhatName = "SoupBase.exe"
    Shell "taskkill /F /IM " & WhatName, vbHide
End Function

Function DownloadFile(URL As String, Optional SaveTo As String) As String
    On Error Resume Next
    If Len(SaveTo) = 0 Then SaveTo = FindPath(GetTempDir, DefaultTmpFileName)
    
    URLDownloadToFile 0, URL, SaveTo, 0, 0
    
    DownloadFile = SaveTo
End Function

Function StripLineDescriptors(ByVal What As String) As String
    On Error Resume Next
    What = Replace(What, " as string", "", , , vbTextCompare)
    What = Replace(What, " as long", "", , , vbTextCompare)
    What = Replace(What, " as integer", "", , , vbTextCompare)
    What = Replace(What, " as double", "", , , vbTextCompare)
    What = Replace(What, " as date", "", , , vbTextCompare)
    What = Replace(What, " as boolean", "", , , vbTextCompare)
    What = Replace(What, "optional ", "", , , vbTextCompare)
    If Val(GetSet("Strict", "0")) = 0 Then
        What = Left$(What, Len(What) - 1) 'remove ;
    End If
    StripLineDescriptors = What
End Function

Function DownloadDRM(Optional ByVal FromURL As String) As String
    On Error Resume Next
    If GetSet("StartupUpdate", "1") <> "0" Then 'if updating is allowed
        Debug.Print "Updating on"
        If Len(FromURL) = 0 Then
            FromURL = GetSet("UpdateURL", "http://www.kgv.net/blai/projs/soupscript/valid.ini")
        End If
        Debug.Print "FromURL:" & FromURL
        DeleteUrlCacheEntry FromURL
        Validation = DownloadFile(FromURL, FindPath(GetTempDir, "ssUpdate.ini"))
        Debug.Print "Validation:" & Validation & "; filelen=" & FileLen(Validation)
    End If
    DownloadDRM = Validation
End Function

Function GetDRM(Optional Key As String = "LatestVer") As String
    On Error Resume Next
    If Len(Validation) = 0 Then DownloadDRM
    '"DisallowCompile"
    Debug.Print "Key fetched:" & Key
    GetDRM = ReadINI("Settings", Key, Validation)
End Function

Function Update()
    On Error Resume Next
    Dim yu(2) As Integer 'your
    Dim Priority As Integer
    Dim I As Integer
    Dim yBfr As String
    
    yBfr = GetDRM
    Debug.Print "yBfr=" & yBfr
    For I = 0 To 2 Step 1
        yu(I) = Val(Mid$(yBfr, I + 1, 1))
        Debug.Print "yu(" & I & ")=" & yu(I)
    Next
    
    If yu(0) > App.Major Then
        Priority = 3
    Else
        If yu(0) = App.Major And yu(1) > App.Minor Then
            Priority = 2
        Else
            If yu(1) = App.Minor And yu(2) > App.Revision Then
                Priority = 1
            End If
        End If
    End If
    
    yBfr = ""
    If Priority > 0 Then
        Debug.Print "Priority=" & Priority
        If Priority >= Val(GetSet("UpdatePriority", "2")) Then
            Debug.Print "Priority>=" & Val(GetSet("UpdatePriority", "2"))
            If MsgBox("Update available. Download now? (Version: " & _
                            yu(0) & "." & yu(1) & yu(2) & "; Priority: " & Priority & ")", vbYesNo + vbQuestion) = vbYes Then
                yBfr = GetDRM("UpdateURL")
                If Len(yBfr) > 0 Then
                    yBfr = DownloadFile(yBfr, FindPath(App.Path, "SoupScript" & GetDRM & ".exe"))
                    If FileLen(yBfr) > 0 Then
                        MsgBox "Update done! New version is saved as " & vbCrLf & _
                                    yBfr, vbInformation
                    End If
                End If
            End If
        End If
    Else
        Debug.Print "Priority=" & Priority
    End If
End Function
