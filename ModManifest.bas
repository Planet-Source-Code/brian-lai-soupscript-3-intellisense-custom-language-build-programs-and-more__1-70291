Attribute VB_Name = "ModManifest"
'[ModManifest]
Option Explicit

Private Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type

Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long

Private Const SEM_NOGPFAULTERRORBOX = &H2&
Private m_bInIDE As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public Function XPVB()
    On Error Resume Next
    Dim iccex As tagInitCommonControlsEx
    With iccex
        .lngSize = LenB(iccex)
        .lngICC = ICC_USEREX_CLASSES
    End With
    InitCommonControlsEx iccex
    On Error GoTo 0
End Function

Public Function FindPath(Parent As String, Optional Child As String, Optional Divider As String = "\") As String
    On Error Resume Next
    If Right$(Parent, 1) = Divider Then Parent = Left$(Parent, Len(Parent) - 1)
    If Left$(Child, 1) = Divider Then Child = Mid$(Child, 2)
    FindPath = Parent & Divider & Child
End Function

Sub Main()
    On Error Resume Next
    Dim MUSTRUN As Boolean
    
    Kill FindPath(GetTempDir, "ssUpdate.ini") 'update file must be deleted
    
    XPVB
    If Len(Command$) > 0 Then
        Dim K As String
        K = Command$
        
        If InStr(1, K, "/run", vbTextCompare) > 0 Then
            MUSTRUN = True
            K = Replace(K, "/run", "", , , vbTextCompare)
        End If
        
        If Left$(K, 1) = """" Then K = Right$(K, Len(K) - 1)
        If Right$(K, 1) = """" Then K = Left$(K, Len(K) - 1) 'trimming out quotes
        frmMain.LoadFile K
        If MUSTRUN Then frmMain.RunCode
    Else
        'frmMain.titFileNew_Click
    End If
    frmMain.Show
End Sub

Public Function IsIDE() As Boolean
    On Error Resume Next
    IsIDE = (App.LogMode = 0)
End Function

Public Sub UnloadApp()
    On Error Resume Next
    If Not IsIDE() Then
        SetErrorMode SEM_NOGPFAULTERRORBOX
        LoadLibrary "comctl32.dll"
    End If
End Sub

