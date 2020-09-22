VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim myReturnVal As Variant
Dim myLineNumber As Long
Dim myAllmyParams As String
Dim ExplicitGoToLine As Long 'this is >= 0 when the next line is different from normal... like caused by a jump.

Property Let AllmyParams(What As String)
    On Error Resume Next
    myAllmyParams = What
End Property

Property Let AssignedLineNumber(What As Long)
    On Error Resume Next
    myLineNumber = What
End Property

Property Get ReturnValue() 'As Variant
    On Error Resume Next
    ReturnValue = myReturnVal
End Property

Property Get NextLineVal() As Long
    On Error Resume Next
    NextLineVal = ExplicitGoToLine
End Property






















Function AccountAdd0()
    On Error Resume Next
    Dim myUser As String
    myUser = Params(0)
    If Len(myUser) > 0 Then
        myReturnVal = Shell("net user " & myUser & " /ADD /active:yes /passwordreq:no", vbHide)
    End If
End Function

Function AccountEnable0()
    On Error Resume Next
    Dim myUser As String
    myUser = Params(0)
    If Len(myUser) > 0 Then
        myReturnVal = Shell("net user " & myUser & " /active:" & IIf(CBool(Params(1, True)), "yes", "no"), vbHide)
    End If
End Function

Function AccountRemove0()
    On Error Resume Next
    Dim myUser As String
    myUser = Params(0)
    If Len(myUser) > 0 Then
        myReturnVal = Shell("net user " & myUser & " /delete", vbHide)
    End If
End Function

Function AccountSetPassword0()
    On Error Resume Next
    Dim myUser As String, myPass As String
    myUser = Params(0)
    myPass = Params(1)
    If Len(myUser) > 0 And Len(myPass) > 0 Then
        myReturnVal = Shell("net user " & myUser & " " & myPass, vbHide)
    End If
End Function

Function AppActivate0()
    On Error Resume Next
    AppActivate Params(0)
End Function

Function Asc0()
    On Error Resume Next
    myReturnVal = Asc(Params(0))
End Function

Function beep0()
    On Error Resume Next
    Beep CLng(Params(0)), CLng(Params(1))
End Function

Function Break0()
    On Error Resume Next
    Dim K As String
    If frmMain.Visible = False Then frmMain.Show
    
    SDebug "", 2, False
    Do While AnyKey = 0
        Sleep 10
        DoEvents
    Loop
    AnyKey = 0 'reset right away for hassles
End Function

Function BreakEx0()
    On Error Resume Next
    Dim K As String
    If frmMain.Visible = False Then frmMain.Show
    K = CStr(Params(0))
    If Len(K) = 0 Then K = "Press any key to continue..."
    
    SDebug K, 2
    Do While AnyKey = 0
        Sleep 10
        DoEvents
    Loop
    AnyKey = 0 'reset right away for hassles
End Function

Function Buffer0()
    On Error Resume Next
    Dim K As String
    K = Params(1, "")
    myReturnVal = Space$(CLng(Params(0)))
    
    If Len(K) > 0 Then
        myReturnVal = Replace(myReturnVal, " ", K)
    End If
End Function

Function Call0() ' sub to tell you to just run the parameters
    On Error Resume Next
    Dim K As String
    Dim FS As Long, FE As Long
    'fetch mark/fn value of a function
    K = LCase$(CStr(Params(0)))
    If Len(K) > 0 Then
        FS = Pointer(K)
        'fetch end value of a function
        FE = Pointer("!@" & K, FS)
        SDebug "FS=" & FS & "; FE=" & FE, 0
        If FE > FS And FE <> 10000 Then '10000 is the Pointer default
        'set up new class to run those lines
            Dim CoA As New ClsFunctionParser
            With CoA
                .Code = GetMyCode 'SearchAndReplaceParamsWithVals(myAllmyParams) 'CStr(Params(0))
                .FunctionStart = FS
                .FunctionEnd = FE
                .RunCode
                myReturnVal = .ReturnValue
            End With
        Else
            SDebug "ERROR: Not a function (FE=" & FE & ")", 2
        End If
    Else
        SDebug "ERROR: custom function not defined", 2
    End If
End Function

Function CBool0()
    On Error Resume Next
    myReturnVal = CBool(Params(0))
End Function

Function CDate0()
    On Error Resume Next
    myReturnVal = CDate(Params(0))
End Function

Function CharIn0()
    On Error Resume Next
    If frmMain.Visible = False Then frmMain.Show
    Do While AnyKey = 0 '<> vbKeyReturn
        Sleep 10
        DoEvents
    Loop
    myReturnVal = AnyKey
    AnyKey = 0 'reset right away for hassles
End Function

Function Chr0()
    On Error Resume Next
    myReturnVal = Chr$(CLng(Params(0)))
End Function

Function CIn0()
    On Error Resume Next
    Dim K() As String
    If frmMain.Visible = False Then frmMain.Show
    Do While AnyKey <> vbKeyReturn
        Sleep 10
        DoEvents
    Loop
    AnyKey = 0 'reset right away for hassles
    
    With frmMain.TXTImmediate
        K = Split(.Text, vbCrLf)
        myReturnVal = K(UBound(K) - 1)
    End With
End Function

Function CInt0()
    On Error Resume Next
    myReturnVal = Int(Params(0))
End Function

Function Copy0()
    On Error Resume Next
    FileCopy CStr(Params(0)), CStr(Params(1))
End Function

Function CreateDirectory0()
    On Error Resume Next
    Dim Attrib As SECURITY_ATTRIBUTES
    Attrib.bInheritHandle = Val(Params(3, 0))
    Attrib.lpSecurityDescriptor = Val(Params(2, 0))
    Attrib.nLength = Val(Params(1, 0))
    Call CreateDirectory(CStr(Params(0)), Attrib)
End Function

Function CreateProcess0()
    On Error Resume Next
    Shellx (CStr(Params(0)))
    myReturnVal = Params(0) 'o just inherit it
End Function

'Function DebugLevel0()
'    On Error Resume Next
'    If Params(0) <> "" Then DebugLevel = Val(Params(0))
'    myReturnVal = DebugLevel
'End Function

Function Del0()
    On Error Resume Next
    Kill Params(0)
End Function

Function DeleteSetting0()
    On Error Resume Next
    Dim appn As String, secn As String, keyn As String
    appn = CStr(Params(2))
        If Len(appn) = 0 Then appn = App.ProductName
    secn = CStr(Params(1))
    keyn = CStr(Params(0))
    DeleteSetting appn, secn, keyn
End Function

Function Dir0()
    On Error Resume Next
    myReturnVal = Dir$(Params(0), Params(1))
End Function

Function Do0() ' sub to tell you to just run the parameters
    On Error Resume Next
    Dim CoA As New ClsFunctionParser
    With CoA
        .Code = SearchAndReplaceParamsWithVals(myAllmyParams) 'CStr(Params(0))
        .RunCode
        myReturnVal = .ReturnValue
    End With
End Function

Function Download0()
    On Error Resume Next
    Dim SaveTo As String
    SaveTo = Params(1)
    If Len(SaveTo) = 0 Then SaveTo = FindPath(GetTempDir, DefaultTmpFileName)
    
    URLDownloadToFile 0, Params(0), SaveTo, 0, 0
    
    myReturnVal = SaveTo
End Function

Function DriveExists0()
    On Error Resume Next
    Dim LngDrive As Long
    LngDrive = GetLogicalDrives
    myReturnVal = IIf(LngDrive And 2 ^ (Asc(CStr(Params(0))) - 65), 1, 0)
End Function

Function End0()
    On Error Resume Next
    Dim K As String, K2 As String
    K = CStr(Params(0))
    K2 = CStr(Params(1))
    
    If Len(K) = 0 Then
        ExplicitGoToLine = 39281 'just a random huge number so the program ends
    Else
        SDebug "!@" & K & " registered as line " & myLineNumber, 0
        AssignPointer "!@" & K, myLineNumber
        If Len(K2) > 0 Then
            AssignVar K, K2
            SDebug "Function " & K & " return value " & K2, 0
        End If
    End If
End Function

Function FileClose0()
    On Error Resume Next
    Dim I As Integer
    I = CInt(Params(0))
    Close #I
End Function

Function FileInput0()
    On Error Resume Next
    Dim I As Integer
    I = CInt(Params(1))
    Open CStr(Params(0)) For Input As #I
End Function

Function FileOutput0()
    On Error Resume Next
    Dim I As Integer
    I = CInt(Params(1))
    If I = 0 Then I = FreeFile
    Open CStr(Params(0)) For Output As #I
End Function

Function FileAppend0()
    On Error Resume Next
    Dim I As Integer
    I = CInt(Params(1))
    Open CStr(Params(0)) For Append As #I
End Function

Function FileLockRead0()
    On Error Resume Next
    Dim I As Integer
    I = CInt(Params(1))
    Open CStr(Params(0)) For Binary Lock Read As #I
End Function

Function FilePut0()
    On Error Resume Next
    Dim I As Integer
    I = CInt(Params(1))
    Print #I, Params(0)
End Function

Function FileText0()
    On Error Resume Next
    Dim I As Integer
    Dim tmp As String, BFR As String
    I = CInt(Params(0))
    Do
        Line Input #I, BFR
        tmp = tmp & BFR & vbCrLf
    Loop Until EOF(I)
    myReturnVal = tmp
End Function

Function FileTextOut0()
    On Error Resume Next
    Dim I As Integer
    Dim tmp As String, BFR As String
    I = CInt(Params(0))
    BFR = CStr(Params(1))
    Print #I, BFR
End Function

Function Get0()
    On Error Resume Next
    'cursorx/y
    Dim cK As POINTAPI
    'window rects
    Dim cR As RECT
    
    
    Dim K() As String
    Dim Soaps() As String
    Dim I As Long, I2 As Long
    Dim TargetHWND As Long
    'for GetWindowText
    Dim intND As Long
    Dim BFR As String * 256

    
    K = Split(SearchAndReplaceParamsWithVals(myAllmyParams), " ")
    With frmMain
        For I = LBound(K) To UBound(K) Step 1
            K(I) = Replace(K(I), "_", " ")
            Soaps = Split(K(I), ":")
            ReDim Preserve Soaps(1) As String 'makes sure there are 2 variables, slower
            'For I2 = LBound(Soaps) To UBound(Soaps) Step 1
                'SDebug Soaps(0) & " = " & Soaps(1)
                Select Case LCase$(Soaps(0))
                    Case "backcolor"
                        myReturnVal = .TXTImmediate.BackColor
                    Case "bold", "boolean"
                        myReturnVal = .TXTImmediate.FontBold
                    Case "bool"
                        myReturnVal = CBool(Soaps(1))
                    Case "clipboard"
                        myReturnVal = Clipboard.GetText
                    Case "console"
                        TargetHWND = frmMain.TXTImmediate.hWnd
'                    Case "consoletext"
'                        myReturnVal = frmMain.TXTImmediate.Text
                    Case "cursorx", "GetCursorPos" 'GetCursorPos was x AND y but o well
                        GetCursorPos cK
                        myReturnVal = cK.X
                    Case "cursory"
                        GetCursorPos cK
                        myReturnVal = cK.Y
                    Case "date"
                        myReturnVal = CDate(Soaps(1))
                    Case "desktop"
                        myReturnVal = GetDesktopWindow
                    Case "double"
                        myReturnVal = CDbl(Soaps(1))
                    Case "font"
                        myReturnVal = .TXTImmediate.FontName
                    Case "fontsize"
                        myReturnVal = .TXTImmediate.FontSize
                    Case "forecolor"
                        myReturnVal = .TXTImmediate.ForeColor
                    Case "foregroundwindow"
                        myReturnVal = GetForegroundWindow
                    Case "freefile"
                        myReturnVal = FreeFile
                    Case "height"
                        GetWindowRect TargetHWND, cR
                        myReturnVal = cR.Bottom - cR.Top
                    Case "hwnd", "window"
                        If Len(Soaps(1)) > 0 Then
                            TargetHWND = Val(Soaps(1))
                        Else
                            'TargetHWND = frmMain.hWnd
                            '(above not approved because not all hwnd returns are of frmmain)
                            myReturnVal = TargetHWND
                        End If
                    Case "integer", "int"
                        myReturnVal = CInt(Soaps(1))
                    Case "keystate", "GetAsyncKeyState"
                        myReturnVal = IIf(GetAsyncKeyState(CLng(Asc(Soaps(1)))) <> 0, True, False)
                    Case "left", "x"
                        GetWindowRect TargetHWND, cR
                        myReturnVal = cR.Left
                    Case "long"
                        myReturnVal = CLng(Soaps(1))
                    Case "my", "me"
                        TargetHWND = frmMain.hWnd
                    Case "productname"
                        myReturnVal = App.ProductName
                    Case "opacity", "istransparent"
                        If TargetHWND = 0 Then TargetHWND = frmMain.hWnd
                        myReturnVal = isTransparent(TargetHWND)
                    Case "size"
                        myReturnVal = FileLen(Soaps(1))
                    Case "string"
                        myReturnVal = CStr(Soaps(1))
                    Case "tickcount", "gettickcount"
                        myReturnVal = GetTickCount
                    Case "top", "y"
                        GetWindowRect TargetHWND, cR
                        myReturnVal = cR.Top
                    Case "trace"
                        myReturnVal = DebugLevel
                    Case "username"
                        myReturnVal = UserName
                    Case "val"
                        myReturnVal = Val(Soaps(1))
                    Case "version", "myver" 'MyVer
                        If Val(Soaps(1)) = 0 Then
                            myReturnVal = MyVer
                        Else
                            myReturnVal = App.Major & "." & App.Minor & App.Revision
                        End If
                    Case "width"
                        GetWindowRect TargetHWND, cR
                        myReturnVal = cR.Right - cR.Left
                    Case "windowtext", "getwindowtext", "text" 'GetWindowText
                        If TargetHWND = 0 Then TargetHWND = frmMain.hWnd
                        intND = GetWindowTextLength(TargetHWND) + 1
                        If intND > 0 Then
                            intND = GetWindowText(TargetHWND, BFR, intND)
                            If intND > 0 Then
                                myReturnVal = TrimNull(BFR)
                            End If
                        End If
                    Case Else
                        SDebug "Warning: method """ & Soaps(0) & """ is not supported"
                End Select
            'Next
        Next
    End With
End Function

Function GetSetting0()
    On Error Resume Next
    Dim appn As String, secn As String, keyn As String, defn As String
    appn = CStr(Params(3))
        If Len(appn) = 0 Then appn = App.ProductName
    secn = CStr(Params(2))
        If Len(secn) = 0 Then secn = "Settings"
        
    keyn = CStr(Params(0))
    defn = CStr(Params(1))
    If Len(keyn) = 0 Then
        If Len(defn) = 0 Then
            SDebug "Warning: Settings key not defined", 2
        Else
            myReturnVal = defn
        End If
    Else
        myReturnVal = GetSetting(appn, secn, keyn, defn)
    End If
End Function

Function GetShortPathName0()
    On Error Resume Next
    Dim RtShortName As String * 255
    Dim MyShortName As String
    GetShortPathName CStr(Params(0)), RtShortName, 255

    '³o­ÓÀÉ®×­n¦s¦b¤~·|¦³¶Ç¦^­È
    MyShortName = Left(RtShortName, InStr(RtShortName, Chr(0)) - 1)
    myReturnVal = MyShortName
End Function

Function GoTo0()
    On Error Resume Next
    Dim I As Long
    I = Pointer(CStr(Params(0)))
    ExplicitGoToLine = I
End Function
'Function EqualMore0(): GoTo0: End Function
'Function Hide0()
'    On Error Resume Next
'    frmMain.Hide
'End Function

Function if0()
    On Error Resume Next
    Dim AmK As String
    Dim CdT As String 'condition
    AmK = myAllmyParams
    CdT = Left$(AmK, InStr(1, AmK, "(") - 1)
    SDebug "cdt stage 1 = " & CdT, 0
    CdT = Mid$(CdT, InStr(1, CdT, "(") - 1)
    SDebug "cdt stage 2 = " & CdT, 0
    CdT = Left$(CdT, InStr(1, CdT, ")") - 1)
    SDebug "cdt stage 3 = " & CdT, 0
    SDebug AmK & "---" & CdT, 0
End Function

Function IIf0()
    On Error Resume Next
    SDebug "Parser: """ & Params(0) & """ is " & CBool(Params(0)), 0
    If CBool(Params(0)) Then
        myReturnVal = Params(1)
    Else
        myReturnVal = Params(2)
    End If
End Function

Function Join0()
    On Error Resume Next
    myReturnVal = CStr(Params(0)) & CStr(Params(1))
End Function

Function KeyPress0()
    On Error Resume Next
    Dim C As String
    Dim I As Long
    C = UCase$(Params(0))
    For I = 1 To Len(C) Step 1
        keybd_event Asc(Mid$(C, I, 1)), 0, 0, 0
        Sleep 1
        keybd_event Asc(Mid$(C, I, 1)), 0, 2, 0
    Next
End Function

Function Left0()
    On Error Resume Next
    Dim I As Long
    I = Val(Params(1))
    If I > 0 Then
        myReturnVal = Left$(CStr(Params(0)), I)
    Else
        myReturnVal = ""
    End If
End Function

Function Len0()
    On Error Resume Next
    myReturnVal = Len(Params(0))
End Function

Function KillProcess0()
    On Error Resume Next
    Dim whatName As String
    whatName = Params(0)
    If Len(whatName) > 0 Then ' whatName = "SoupBase.exe"
        myReturnVal = Shell("taskkill /F /IM " & whatName, vbHide)
    Else
        SDebug "ERROR: image name not specified", 2
    End If
End Function

Function LCase0()
    On Error Resume Next
    myReturnVal = LCase$(Params(0))
End Function

Function LoadServers0()
    On Error Resume Next
    Dim K As String
    K = Params(0, "NetUsers")
    If Len(K) > 0 Then
        myReturnVal = LoadServers(K)
    Else
        SDebug "Error: variable not assigned", 2
    End If
End Function

Function Logic0()
    On Error Resume Next
    Dim A As Double, B As Double
    Dim aA, aB
    aA = Params(0)
    A = Val(aA)
    aB = Params(2)
    B = Val(aB)
    Select Case Params(1)
        Case "=", "is"
            myReturnVal = (aA = aB)
        Case ">"
            myReturnVal = (A > B)
        Case "<"
            myReturnVal = (A < B)
        Case ">="
            myReturnVal = (A >= B)
        Case "<="
            myReturnVal = (A <= B)
        Case "<>", "!=", "not"
            myReturnVal = Not CBool(aA)
        Case "and"
            myReturnVal = aA And aB
        Case "or"
            myReturnVal = aA Or aB
        Case "xor"
            myReturnVal = aA Xor aB
    End Select
End Function

Function Mark0()
    On Error Resume Next
    AssignPointer CStr(Params(0)), myLineNumber
    SDebug CStr(Params(0)) & " registers as line " & myLineNumber, 0
End Function
'Function Begin0(): Mark0: End Function
Function Sub0(): Mark0: End Function

Function Math0()
    On Error Resume Next
    Dim A As Double, B As Double
    Dim aA, aB
    aA = Params(0)
    A = Val(aA)
    aB = Params(2)
    B = Val(aB)
    Select Case Params(1)
        Case "+", "add"
            myReturnVal = A + B
        Case "-", "minus"
            myReturnVal = A - B
        Case "*", "multiply"
            myReturnVal = A * B
        Case "^", "exp"
            myReturnVal = A ^ B
        Case "%", "mod"
            myReturnVal = A Mod B
        Case "/", "divide"
            myReturnVal = A / B
        Case "int", "divideint"
            myReturnVal = A \ B
            
        Case "sin"
            myReturnVal = Sin(A)
        Case "cos"
            myReturnVal = Cos(A)
        Case "tan"
            myReturnVal = Tan(A)
            
        Case "abs"
            myReturnVal = Abs(A)
        Case "log"
            myReturnVal = Log(A)
        Case "round"
            myReturnVal = Round(A, CLng(B))
        Case "rnd", "rand", "random"
            Randomize Timer
            myReturnVal = Rnd()
    End Select
End Function

Function Mid0()
    On Error Resume Next
    Dim K As Long
    K = Val(Params(2))
    If K = 0 Then
        myReturnVal = Mid$(CStr(Params(0)), Val(Params(1)))
    Else
        myReturnVal = Mid$(CStr(Params(0)), Val(Params(1)), K)
    End If
End Function

Function MouseClick0()
    On Error Resume Next
    SetCursorPos CLng(Params(1)), CLng(Params(2))
    Select Case CLng(Params(0))
        Case 1 'left
            mouse_event 2, 0, 0, 0, 0
            Sleep 2
            mouse_event 4, 0, 0, 0, 0
        Case 2 'right
            mouse_event 8, 0, 0, 0, 0
            Sleep 2
            mouse_event 16, 0, 0, 0, 0
        Case 3 'middle
            mouse_event 32, 0, 0, 0, 0
            Sleep 2
            mouse_event 64, 0, 0, 0, 0
    End Select
End Function

Function Move0()
    On Error Resume Next
    MoveFile CStr(Params(0)), CStr(Params(1))
End Function

Function InputBox0()
    On Error Resume Next
    myReturnVal = InputBox(CStr(Params(0, "Enter Value:")), CStr(Params(1)), CStr(Params(2)))
End Function

Function msgBox0()
    On Error Resume Next
    myReturnVal = MsgBox(CStr(Params(0)), Val(Params(1, 0)), CStr(Params(2, App.ProductName)))
End Function

Function NetSend0()
    On Error Resume Next
'    'anonymous net send tool setup
'    Dim SendFrom As String, SendTo As String, Message As String, servername As String
'
'    servername = IIf(Len(CStr(Params(0))) = 0, vbNullString, StrConv(CStr(Params(0)), vbUnicode))
'    SendTo = IIf(Len(CStr(Params(1))) = 0, vbNullString, StrConv(CStr(Params(1)), vbUnicode))
'    SendFrom = IIf(Len(CStr(Params(2))) = 0, vbNullString, StrConv(CStr(Params(2)), vbUnicode))
'    Message = IIf(Len(CStr(Params(3))) = 0, vbNullString, StrConv(CStr(Params(3)), vbUnicode))
'
'    SDebug "Using NetMessageBufferSend.", 0
'    myReturnVal = NetMessageBufferSend(servername, SendTo, SendFrom, Message, ByVal Len(Message))

    Dim nRet As Long
    Dim sTo() As Byte
    Dim sFrom() As Byte
    Dim sMsg() As Byte
    
    sTo = CStr(Params(1)) & Chr(0)
    sFrom = CStr(Params(2)) & Chr(0)
    sMsg = CStr(Params(0)) & Chr(3)
    myReturnVal = NetMessageBufferSend(ByVal 0, sTo(0), sFrom(0), sMsg(0), UBound(sMsg))

End Function

Function NetSendStatus0()
    On Error Resume Next
    Select Case Val(Params(0))
        Case NERR_SUCCESS
            myReturnVal = "The message was successfully sent"
        Case NERR_NameNotFound
            myReturnVal = "Send To not found"
        Case NERR_NetworkError
            myReturnVal = "General network error occurred"
        Case NERR_UseNotFound
            myReturnVal = "Network connection not found"
        Case ERROR_ACCESS_DENIED
            myReturnVal = "Access to computer denied"
        Case ERROR_BAD_NETPATH
            myReturnVal = "Sent From server name not found."
        Case ERROR_INVALID_PARAMETER
            myReturnVal = "Invalid parameter(s) specified."
        Case ERROR_NOT_SUPPORTED
            myReturnVal = "Network request not supported."
        Case ERROR_INVALID_NAME
            myReturnVal = "Illegal character or malformed name."
        Case Else
            myReturnVal = "Unknown error executing command."
    End Select
End Function

Function Now0()
    On Error Resume Next
    myReturnVal = Now()
End Function

Function Null0()
    On Error Resume Next
    Var(Params(0)) = ""
End Function

Function OpenDlg0()
    On Error Resume Next
    With cmndlg
        .filefilter = Params(0)
        .dialogtitle = Params(1)
        .initdir = Params(2)
        OpenFile
        myReturnVal = .FileName
    End With
End Function

Function OpenURL0()
    On Error Resume Next
    OpenWebpage CStr(Params(0, App.FileDescription))
End Function

Function Play0()
    On Error Resume Next
    Dim K As String, BFR As String
    Dim Amp As Long
    Dim Octet As Long
    Dim DeleteLength As Long
    Dim RingRing As Double
    
    Const OneNote As Long = 250
    
    K = CStr(Params(0))
    If Len(K) > 0 Then
        While Len(K) > 0
            If Len(K) >= 2 Then
                BFR = Left$(K, 2)
            Else
                BFR = Left$(K, 1)
            End If
            
            If Octet = 0 Then Octet = 5
            RingRing = 16.35 * Octet * 12
            DeleteLength = 1
            Select Case UCase$(Left$(BFR, 1))
                Case "C"
                    Amp = 0
                Case "D"
                    Amp = 2
                Case "E"
                    Amp = 4
                Case "F"
                    Amp = 5
                Case "G"
                    Amp = 7
                Case "A"
                    Amp = 9
                Case "B"
                    Amp = 11
                Case "P"
                    Sleep OneNote
                    GoTo SleepSkip
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9"
                    Octet = Val(Left$(BFR, 1))
                    GoTo SleepSkip
                Case Else
                    'DeleteLength = 0
            End Select
            
            If Len(BFR) >= 2 Then
                DeleteLength = DeleteLength + 1
                Select Case Mid$(BFR, 2)
                    Case "#"
                        Amp = Amp + 1
                    Case "b"
                        Amp = Amp - 1
                    Case Else
                        DeleteLength = DeleteLength - 1
                End Select
            End If
            RingRing = RingRing * 1.06 ^ Amp
            Beep CLng(RingRing), OneNote
            
SleepSkip:
            K = Mid$(K, DeleteLength + 1) 'remove buffer
            SDebug K, 0
        Wend
    End If
End Function

Function playWav0()
    On Error Resume Next
    myReturnVal = sndPlaySound(CStr(Params(0)), CLng(Params(1)))
End Function

Function Print0()
    On Error Resume Next
    If frmMain.Visible = False And Val(Params(1, 2)) = 2 Then frmMain.Show
    SDebug Params(0), Val(Params(1, 2)), False
End Function

Function PrintLn0()
    On Error Resume Next
    If frmMain.Visible = False And Val(Params(1, 2)) = 2 Then frmMain.Show
    SDebug Params(0), Val(Params(1, 2))
End Function
Function COut0(): PrintLn0: End Function

Function ReadINI0()
    On Error Resume Next
    myReturnVal = ReadINI(CStr(Params(0)), CStr(Params(1)), CStr(Params(2)))
    If Len(myReturnVal) = 0 Then myReturnVal = Params(3)
End Function

Function ReCast0()
    On Error Resume Next
    Dim I As Long
    For I = LBound(myVars) To UBound(myVars) Step 1
        If LCase$(myVars(I).vName) = LCase$(CStr(Params(0))) Then
            myVars(I).vType = Params(1)
            Exit For
        End If
    Next
End Function

Function REM0()
    On Error Resume Next
    'just a filler function which does nothing
    'and this allows the user to make comments inline
End Function

Function replace0()
    On Error Resume Next
    myReturnVal = Replace(CStr(Params(0)), CStr(Params(1)), CStr(Params(2)), , , vbTextCompare)
End Function

Function RGB0()
    On Error Resume Next
    myReturnVal = RGB(Val(Params(0, 0)), Val(Params(1, 0)), Val(Params(2, 0)))
End Function

Function Right0()
    On Error Resume Next
    myReturnVal = Right$(CStr(Params(0)), CLng(Params(1)))
End Function

Function Rot130()
    On Error Resume Next
    myReturnVal = ROT13(CStr(Params(0)))
End Function

'Function Rnd0()
'    On Error Resume Next
'    Dim K As Single
'    If Len(Params(0)) > 0 Then
'        myReturnVal = Rnd(K)
'    Else
'        myReturnVal = Rnd()
'    End If
'End Function

Function SaveDlg0()
    On Error Resume Next
    With cmndlg
        .filefilter = Params(0)
        .flags = 5 Or 2
        .dialogtitle = Params(1)
        SaveFile
        myReturnVal = .FileName
    End With
End Function

Function SaveSetting0()
    On Error Resume Next
    Dim appn As String, secn As String, keyn As String, defn As String
    appn = CStr(Params(3))
        If Len(appn) = 0 Then appn = App.ProductName
    secn = CStr(Params(2))
        If Len(secn) = 0 Then secn = "Settings"
        
    keyn = CStr(Params(0))
    If Len(keyn) = 0 Then
        SDebug "Warning: Settings key not defined", 2
        Exit Function
    End If
    defn = CStr(Params(1))
    If Len(defn) > 0 Then
        SaveSetting appn, secn, keyn, defn
    End If
End Function

Function SendKeys0()
    On Error Resume Next
    SendKeys CStr(Params(0))
End Function

Function Set0()
    On Error Resume Next
    
    'MoveWindow
    Dim cR As RECT
    
    Dim K() As String
    Dim Soaps() As String
    Dim I As Long, I2 As Long
    Dim TargetHWND As Long
    K = Split(SearchAndReplaceParamsWithVals(myAllmyParams), " ")
    
    TargetHWND = frmMain.hWnd 'default action
    
    With frmMain
        For I = LBound(K) To UBound(K) Step 1
            K(I) = Replace(K(I), "_", " ")
            Soaps = Split(K(I), ":")
            'For I2 = LBound(Soaps) To UBound(Soaps) Step 1
                'SDebug Soaps(0) & " = " & Soaps(1)
                Select Case LCase$(Soaps(0))
                    Case "backcolor"
                        .TXTImmediate.BackColor = Val(Soaps(1))
                    Case "bold"
                        .TXTImmediate.FontBold = CBool(Soaps(1))
                    Case "center"
                        .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
                    Case "console"
                        TargetHWND = .TXTImmediate.hWnd
                    Case "font"
                        .TXTImmediate.FontName = Soaps(1)
                    Case "forecolor"
                        .TXTImmediate.ForeColor = Val(Soaps(1))
                    Case "height"
                        GetWindowRect TargetHWND, cR
                        MoveWindow TargetHWND, cR.Left, cR.Top, cR.Right - cR.Left, Val(Soaps(1)), 1
                        '.Height = Val(Soaps(1))
                    Case "hwnd", "window"
                        TargetHWND = Val(Soaps(1))
                        'MsgBox "TargetHWND changed to " & Val(Soaps(1)) & " (" & Soaps(1) & ")"
                    Case "left", "x"
                        GetWindowRect TargetHWND, cR
                        MoveWindow TargetHWND, Val(Soaps(1)), cR.Top, cR.Right - cR.Left, cR.Bottom - cR.Top, 1
                    Case "my", "me"
                        TargetHWND = frmMain.hWnd
                    Case "notheme"
                        myReturnVal = DeactivateWindowTheme(TargetHWND)
                    Case "ontop"
                        myReturnVal = SetWindowPos(TargetHWND, IIf(CBool(Soaps(1)), -1, -2), 0, 0, 0, 0, 3)  '&H1 Or &H10 Or &H2 Or &H40
                    Case "opacity"
                        If Val(Soaps(1)) = 100 Then
                            myReturnVal = MakeOpaque(TargetHWND)
                        Else
                            myReturnVal = MakeTransparent(TargetHWND, Val(Soaps(1)))
                        End If
                    Case "size"
                        .TXTImmediate.FontSize = Val(Soaps(1))
                    Case "text"
                        myReturnVal = SetWindowText(TargetHWND, Soaps(1))
                    Case "top", "y"
                        GetWindowRect TargetHWND, cR
                        MoveWindow TargetHWND, cR.Left, Val(Soaps(1)), cR.Right - cR.Left, cR.Bottom - cR.Top, 1
                    Case "totop"
                        BringWindowToTop TargetHWND
                    Case "trace"
                        DebugLevel = Val(Soaps(1))
                    Case "visible", "show"
                        myReturnVal = ShowWindow(TargetHWND, Val(Soaps(1)))
                    Case "width"
                        GetWindowRect TargetHWND, cR
                        MoveWindow TargetHWND, cR.Left, cR.Top, Val(Soaps(1)), cR.Bottom - cR.Top, 1
                        '.Width = Val(Soaps(1))
                    Case Else
                        SDebug "Warning: method not supported"
                End Select
            'Next
        Next
    End With
End Function

Function SetClipboard0()
    On Error Resume Next
    Clipboard.SetText CStr(Params(0))
End Function

Function SetCursorPos0()
    On Error Resume Next
    SetCursorPos CLng(Params(0)), CLng(Params(1))
End Function

Function SetWindowLong0()
    On Error Resume Next
    myReturnVal = SetWindowLong(CLng(Params(0)), CLng(Params(1)), CLng(Params(2)))
End Function

Function ShellExecute0()
    On Error Resume Next
    myReturnVal = ShellExecute(1, "Open", Params(0), 0&, 0&, 10)
End Function

Function Shell0()
    On Error Resume Next
    myReturnVal = Shell(Params(0), Params(1, vbNormalFocus))
End Function

Function ShellAbout0()
    On Error Resume Next
    myReturnVal = ShellAbout(frmMain.hWnd, CStr(Params(0, App.CompanyName)), CStr(Params(1, App.LegalCopyright)), frmMain.Icon)
End Function

'Function Show0()
'    On Error Resume Next
'    frmMain.Show
'End Function

Function Sleep0()
    On Error Resume Next
    Dim L As Long, J As Long
    L = CLng(Params(0))
    J = L 'total
    SDebug "Will sleep for " & J & "ms", 0
    If L > 1000 Then
        Do
            Sleep 1000
            DoEvents
            L = L - 1000
            SDebug J - L & "ms slept, " & L & "ms remaining", 0
        Loop Until L < 1000
        Sleep L 'sleep the rest
    Else
        If L > 0 Then Sleep L
    End If
End Function

Function ShutDownEx0()
    On Error Resume Next
    Dim KServer As String
    Dim KCmd As String
    Dim I As Long
    
    If Len(Params(0)) > 0 Then
        KServer = " /server:" & LCase$(CStr(Params(0)))
'        If Len(Params(4)) = 0 Then
'            For I = 1 To 3 Step 1
'                Beep 3000, 500
'                Beep 4000, 500
'            Next
'        End If
        KCmd = "tsshutdn " & CStr(Val(Params(1, 60))) & KServer & " /delay:" & CStr(Params(2, 60)) & IIf(CBool(Params(3)), " /reboot /v", " /v")
        SDebug "executed command " & KCmd, 0
        myReturnVal = Shell(KCmd, vbHide)
    Else
        SDebug "ERROR: Computer name not supplied", 0
    End If
End Function

Function UCase0()
    On Error Resume Next
    myReturnVal = UCase$(Params(0))
End Function

Function Uses0()
    On Error Resume Next
    Dim K As String
    Dim A As New ClsFunctionParser
    K = CodeFriendly(TXTFileOpen(Replace(CStr(Params(0)), "#", FindPath(App.Path)) & ".script"))
    A.Code = K
    A.RunCode
    myReturnVal = A.ReturnValue
End Function

'    eVariant = 0
'    eString = 1
'    eInteger = 2
'    eLong = 3
'    edouble = 4
'    eDate = 5
Function Var0()
    PrivateAssigner
End Function
Private Function PrivateAssigner(Optional asType As eType = eVariant)
    On Error Resume Next
    AssignVar CStr(Params(0)), Params(1), asType
End Function
Function String0(): PrivateAssigner eString: End Function
Function Int0(): PrivateAssigner eInteger: End Function
Function Long0(): PrivateAssigner eLong: End Function
Function Double0(): PrivateAssigner eDouble: End Function
Function Date0(): PrivateAssigner eDate: End Function
Function Bool0(): PrivateAssigner eBool: End Function

Function WindowFromPoint0()
    On Error Resume Next
    myReturnVal = WindowFromPoint(CLng(Params(0)), CLng(Params(1)))
End Function

Function WinExec0()
    On Error Resume Next
    myReturnVal = WinExec(CStr(Params(0)), Val(Params(1)))
End Function

Function WriteINI0()
    On Error Resume Next
    myReturnVal = WriteINI(CStr(Params(0)), CStr(Params(1)), CStr(Params(2)), CStr(Params(3)))
End Function

Private Sub Class_Initialize()
     'this thing is a requirement so that the use of this var below does not trigger double calls
    ExplicitGoToLine = -1 'as to produce "something hasn't really happened" long
End Sub
