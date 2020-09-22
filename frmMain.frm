VERSION 5.00
Begin VB.MDIForm frmMain 
   Appearance      =   0  '¥­­±
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   ClientHeight    =   6720
   ClientLeft      =   165
   ClientTop       =   525
   ClientWidth     =   8880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "PFMain"
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.PictureBox picToolbar 
      Align           =   1  '¹ï»ôªí³æ¤W¤è
      Appearance      =   0  '¥­­±
      BorderStyle     =   0  '¨S¦³®Ø½u
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   8880
      TabIndex        =   8
      Top             =   0
      Width           =   8880
      Begin SoupScript.CB btnTB 
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   9
         ToolTipText     =   "Build and Run (F5)"
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   "Run"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   0
         MICON           =   "frmMain.frx":058A
         PICN            =   "frmMain.frx":05A6
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SoupScript.CB btnTB 
         Height          =   375
         Index           =   1
         Left            =   2175
         TabIndex        =   10
         ToolTipText     =   "Stop (F9)"
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   0
         MICON           =   "frmMain.frx":08F8
         PICN            =   "frmMain.frx":0914
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SoupScript.CB btnTB 
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   11
         ToolTipText     =   "Make this script into EXE."
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   "Build..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   0
         MICON           =   "frmMain.frx":0C66
         PICN            =   "frmMain.frx":0C82
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SoupScript.CB btnTB 
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   12
         ToolTipText     =   "New (Ctrl+N)"
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   0
         MICON           =   "frmMain.frx":0FD4
         PICN            =   "frmMain.frx":0FF0
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SoupScript.CB btnTB 
         Height          =   375
         Index           =   4
         Left            =   375
         TabIndex        =   13
         ToolTipText     =   "Open (Ctrl+O)"
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   0
         MICON           =   "frmMain.frx":1342
         PICN            =   "frmMain.frx":135E
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SoupScript.CB btnTB 
         Height          =   375
         Index           =   5
         Left            =   750
         TabIndex        =   14
         ToolTipText     =   "Save (Ctrl+S)"
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   0
         MICON           =   "frmMain.frx":16B0
         PICN            =   "frmMain.frx":16CC
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin SoupScript.Tab Tb 
      Align           =   1  '¹ï»ôªí³æ¤W¤è
      Height          =   300
      Left            =   0
      TabIndex        =   7
      Top             =   375
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   529
      BackColor       =   -2147483633
      CloseButton     =   -1  'True
      BlurForeColor   =   0
      ActiveForeColor =   0
      picture         =   "frmMain.frx":1A1E
      AllTabsForeColor=   -2147483630
      FontName        =   "Tahoma"
   End
   Begin VB.PictureBox picCMDPanel 
      Align           =   4  '¹ï»ôªí³æ¥k¤è
      Appearance      =   0  '¥­­±
      BorderStyle     =   0  '¨S¦³®Ø½u
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5820
      Left            =   8760
      MousePointer    =   9  'ªF-¦è¦V
      ScaleHeight     =   5820
      ScaleWidth      =   120
      TabIndex        =   2
      Top             =   675
      Width           =   120
      Begin VB.PictureBox picCMDHint 
         BackColor       =   &H80000018&
         BorderStyle     =   0  '¨S¦³®Ø½u
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         MousePointer    =   1  '½b¸¹§Îª¬
         ScaleHeight     =   615
         ScaleWidth      =   2775
         TabIndex        =   5
         Top             =   5640
         Width           =   2775
         Begin VB.Label lblCodeHint 
            AutoSize        =   -1  'True
            BackStyle       =   0  '³z©ú
            Caption         =   "Code Help: click on a function"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   2775
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox picDrag 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   135
         Left            =   120
         MousePointer    =   7  '¥_-«n¦V
         ScaleHeight     =   135
         ScaleWidth      =   1695
         TabIndex        =   4
         Top             =   5520
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ListBox lstCMDs 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   5340
         IntegralHeight  =   0   'False
         ItemData        =   "frmMain.frx":1A3A
         Left            =   120
         List            =   "frmMain.frx":1B67
         MousePointer    =   1  '½b¸¹§Îª¬
         OLEDragMode     =   1  '¦Û°Ê
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  '¹ï»ôªí³æ¤U¤è
      BorderStyle     =   0  '¨S¦³®Ø½u
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   8880
      TabIndex        =   0
      Top             =   6495
      Width           =   8880
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Ready"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   30
         TabIndex        =   1
         Top             =   15
         Width           =   465
      End
   End
   Begin VB.Menu titFile 
      Caption         =   "&File"
      Begin VB.Menu titFileNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu titFileOpen 
         Caption         =   "Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu titFileWizard 
         Caption         =   "Wizard..."
      End
      Begin VB.Menu titFileS35890 
         Caption         =   "-"
      End
      Begin VB.Menu titFileSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu titFileSaveAs 
         Caption         =   "Save as..."
      End
      Begin VB.Menu S7852 
         Caption         =   "-"
      End
      Begin VB.Menu titFileClose 
         Caption         =   "Close document"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu titEdit 
      Caption         =   "&Edit"
      Begin VB.Menu titEditArray 
         Caption         =   "Cut"
         Index           =   0
         Shortcut        =   ^X
      End
      Begin VB.Menu titEditArray 
         Caption         =   "Copy"
         Index           =   1
         Shortcut        =   ^C
      End
      Begin VB.Menu titEditArray 
         Caption         =   "Paste"
         Index           =   2
         Shortcut        =   ^V
      End
      Begin VB.Menu titEditArray 
         Caption         =   "Select all"
         Index           =   3
         Shortcut        =   ^A
      End
      Begin VB.Menu titS06 
         Caption         =   "-"
      End
      Begin VB.Menu titFileDeploy 
         Caption         =   "Deploy script"
      End
   End
   Begin VB.Menu titView 
      Caption         =   "&View"
      Begin VB.Menu titViewDebug 
         Caption         =   "Debug Window"
         Shortcut        =   ^G
      End
      Begin VB.Menu titEditCodeBar 
         Caption         =   "Code bar"
         Shortcut        =   ^Q
      End
      Begin VB.Menu S80935 
         Caption         =   "-"
      End
      Begin VB.Menu titViewClr 
         Caption         =   "Colour Scheme"
         Begin VB.Menu titViewClrArray 
            Caption         =   "White (default)"
            Index           =   0
         End
         Begin VB.Menu titViewClrArray 
            Caption         =   "Hacker!"
            Index           =   1
         End
         Begin VB.Menu titViewClrArray 
            Caption         =   "Dark Orange"
            Index           =   2
         End
         Begin VB.Menu titViewClrArray 
            Caption         =   "Ergonomic"
            Index           =   3
         End
      End
      Begin VB.Menu titViewFont 
         Caption         =   "Font..."
      End
   End
   Begin VB.Menu titDebug 
      Caption         =   "&Debug"
      Begin VB.Menu titRun 
         Caption         =   "Run"
         Begin VB.Menu titFileBuild 
            Caption         =   "Build..."
         End
         Begin VB.Menu titRunArray 
            Caption         =   "Build and run"
            Index           =   0
            Shortcut        =   {F5}
         End
         Begin VB.Menu titRunArray 
            Caption         =   "Build and run selected line"
            Enabled         =   0   'False
            Index           =   1
            Shortcut        =   +{F5}
         End
      End
      Begin VB.Menu titS04 
         Caption         =   "-"
      End
      Begin VB.Menu titDebugStop 
         Caption         =   "Stop"
         Shortcut        =   {F9}
      End
      Begin VB.Menu S904 
         Caption         =   "-"
      End
      Begin VB.Menu titDebugBuildOptions 
         Caption         =   "Build options"
         Begin VB.Menu titDebugStrict 
            Caption         =   "Strict Encoding"
            Checked         =   -1  'True
         End
         Begin VB.Menu titDebugRuntime 
            Caption         =   "Runtime DLL"
            Begin VB.Menu titDebugRuntimeArr 
               Caption         =   "Create if not found"
               Checked         =   -1  'True
               Index           =   0
            End
            Begin VB.Menu titDebugRuntimeArr 
               Caption         =   "Always recreate"
               Checked         =   -1  'True
               Index           =   1
            End
            Begin VB.Menu titDebugRuntimeArr 
               Caption         =   "Never create"
               Checked         =   -1  'True
               Index           =   2
            End
            Begin VB.Menu titS05 
               Caption         =   "-"
            End
            Begin VB.Menu titDebugFlush 
               Caption         =   "Delete"
            End
         End
      End
      Begin VB.Menu titDebugNote 
         Caption         =   "Notifications"
         Begin VB.Menu titDebugNoteArray 
            Caption         =   "Show everything"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu titDebugNoteArray 
            Caption         =   "Show warnings and errors (Default)"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu titDebugNoteArray 
            Caption         =   "Show only errors"
            Checked         =   -1  'True
            Index           =   2
         End
      End
   End
   Begin VB.Menu titAbout 
      Caption         =   "&About"
      Begin VB.Menu titAboutProg 
         Caption         =   "About this program"
      End
      Begin VB.Menu titS08 
         Caption         =   "-"
      End
      Begin VB.Menu titAboutHelp 
         Caption         =   "Help"
         Begin VB.Menu titAboutHelpArr 
            Caption         =   "General Help"
            Index           =   0
            Shortcut        =   {F1}
         End
         Begin VB.Menu titAboutHelpArr 
            Caption         =   "What are variables?"
            Index           =   1
         End
         Begin VB.Menu titAboutHelpArr 
            Caption         =   "Special Characters"
            Index           =   2
         End
         Begin VB.Menu titAboutHelpArr 
            Caption         =   "Sample Code"
            Index           =   3
         End
         Begin VB.Menu titAboutHelpArr 
            Caption         =   "Commands reference"
            Index           =   4
         End
      End
      Begin VB.Menu titAboutWeb 
         Caption         =   "Visit Website"
      End
      Begin VB.Menu titAboutUpdates 
         Caption         =   "Updates"
         Begin VB.Menu titAboutUpdateNow 
            Caption         =   "Check now"
         End
         Begin VB.Menu titAboutUpdatesStartup 
            Caption         =   "Check when program starts"
            Checked         =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const Bar0Width As Long = 3000
Dim MyX As Long, MyY As Long 'for use in bar resizing

Private Sub btnTB_Click(Index As Integer)
    'On Error Resume Next
    With btnTB(Index)
        Select Case Index
            Case 0
                titRunArray_Click (0)
            Case 1
                titDebugStop_Click
            Case 2
                titFileBuild_Click
                'PopupMenu titDebugNote, , .Left, .Top + .Height, titDebugNoteArray(1)
            Case 3
                titFileNew_Click
            Case 4
                titFileOpen_Click
            Case 5
                titFileSave_Click
        End Select
    End With
End Sub

Private Sub lblCodeHint_DblClick()
    On Error Resume Next
    Dim K As String, M As String
    If InStr(1, lblCodeHint.Caption, vbCrLf) > 0 Then
        M = Left$(lblCodeHint.Caption, InStr(1, lblCodeHint.Caption, vbCrLf) - 1)
        K = InputBox("Enter new help for " & M & ":" & vbCrLf & vbCrLf & "type - to delete.", , GetSet(M, "", "CodeLib"))
        If Len(K) > 0 Then
            If K <> "-" Then
                SaveSet M, K, "CodeLib"
            Else
                SaveSet M, "", "CodeLib"
            End If
        End If
    End If
End Sub

Private Sub lstCMDs_Click()
    On Error Resume Next
    Dim K As String
    K = lstCMDs.List(lstCMDs.ListIndex)
    If InStr(1, K, "(") > 0 Then
        K = Left$(lstCMDs.List(lstCMDs.ListIndex), InStr(1, lstCMDs.List(lstCMDs.ListIndex), "(") - 1)
    Else
        K = Left$(K, Len(K) - 1) 'removing the ;
    End If
    lblCodeHint.Caption = K & vbCrLf & Replace(GetHint(K), "\n", vbCrLf, , , vbTextCompare)
    picCMDPanel_Resize
End Sub

Private Sub lstCMDs_DblClick()
    On Error Resume Next
    AF.txtCode.SelText = StripLineDescriptors(lstCMDs.List(lstCMDs.ListIndex)) & vbCrLf
    AF.txtCode.SetFocus
End Sub

Private Sub MDIForm_Activate()
    On Error Resume Next
    InitCommonControls
End Sub

Private Sub MDIForm_Load()
    On Error Resume Next
    Me.Caption = App.ProductName & " " & MyVer
    'titEditCodeBar.Checked = False
    ColorizeMe
'    Dim I As Integer
'    I = CInt(Val(GetSet("Verbose", "1")))
'    titDebugNoteArray_Click I
    RefreshMenuCheckBoxes
    
    'PutDLL
    SkinFormEx Me
    
    DoEvents
    If GetSet("StartupUpdate", "1") <> "0" Then Update
    
    Me.Show
    frmWizard.Show 1
End Sub

Function ColorizeMe()
    On Error Resume Next
    With lstCMDs
        .BackColor = CLng(GetSet("BackColor1", CStr(RGB(40, 40, 40))))
        .ForeColor = CLng(GetSet("ForeColor1", CStr(RGB(150, 255, 150))))
        .FontName = GetSet("CodeFont", "Verdana")
        .FontSize = CLng(GetSet("CodeFontSize", "8"))
    End With
End Function

Private Sub MDIForm_Terminate()
    UnloadApp
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    On Error Resume Next
    End 'and I'm cheap!
End Sub

Private Sub picCMDPanel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    picCMDPanel.BackColor = RGB(255, 0, 0)
    MyX = X
End Sub

Private Sub picCMDPanel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim I As Long
    If Button = 1 Then
        I = picCMDPanel.Width - X + MyX
        If I < Bar0 + 600 Then I = Bar0
        If I < Bar0Width + 600 And I > Bar0Width - 600 Then I = Bar0Width
        picCMDPanel.Width = I
    End If
End Sub

Private Sub picCMDPanel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    picCMDPanel.BackColor = &H8000000F
End Sub

Private Sub picCMDPanel_Resize()
    On Error Resume Next
    Dim J As Long, M As Long
    picCMDHint.Move Bar0, picCMDPanel.Height - lblCodeHint.Height, picCMDPanel.Width - Bar0, lblCodeHint.Height
    With picCMDPanel
        If .Width > Bar0 Then
            M = .Width - Bar0
            lstCMDs.Move Bar0, 0, M, picCMDHint.Top '.Height
        End If
    End With
End Sub

Private Sub picDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    MyY = Y
    picDrag.BackColor = RGB(255, 0, 0)
End Sub

Private Sub picDrag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim J As Long
    If Button = 1 Then
        J = Y - MyY
        If J Mod 240 = 0 Then
            J = picDrag.Top + J
            If J >= 300 Then '300 = supposed listbox threshold
                If J <= picCMDPanel.Height - picDrag.Height Then
                    picDrag.Top = J
                    picCMDPanel_Resize
                End If
            End If
        End If
    End If
End Sub

Private Sub picDrag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    picDrag.BackColor = &H8000000F
End Sub

Private Sub Tb_Click(tIndex As Integer)
    On Error Resume Next
    Dim Frm As Form
    For Each Frm In Forms
        If TypeOf Frm Is frmScript Then
            If Frm.Tag = Tb.TabTag(tIndex) Then
                Frm.ZOrder 0
                Exit For
            End If
        End If
    Next
End Sub

Private Sub Tb_DblClick(tIndex As Integer)
    AF.WindowState = IIf(AF.WindowState = 0, 2, 0)
End Sub

Private Sub Tb_TabClose(tIndex As Integer)
    'On Error Resume Next
    Dim K As String: K = Tb.TabTag(tIndex)
    Dim J As Form
    
    For Each J In Forms
        If J.Tag = K Then
            Unload J
            Exit For
        End If
    Next
    
    Set J = AF 'AF remakes form, right? exploit
    
End Sub

Private Sub titAboutHelpArr_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
            MsgBox "Commands are always in this syntax, where square brackets indicate optional item:" & vbCrLf & _
                    vbCrLf & _
                    "[variable=]command[(][parameter][,][parameter][...][)][;]" & vbCrLf & _
                     vbCrLf & _
                     "Examples:" & vbCrLf & _
                     "MsgBox Hello!" & vbCrLf & _
                     "MsgBox Hello!;" & vbCrLf & _
                     "MsgBox Hello!,64,My First Program;" & vbCrLf & _
                     "result=MsgBox(Hello!,64,My First Program);" & vbCrLf & _
                      vbCrLf & _
                      "The last example stores the return of the function to variable 'result'.", vbInformation, "General help"
        Case 1
            OpenWebpage "http://www.kgv.net/blai/projs/soupscript/types.htm#Using_Variables"
        Case 2
            OpenWebpage "http://www.kgv.net/blai/projs/soupscript/types.htm#Special_Characters"
        Case 3
            OpenWebpage "http://www.kgv.net/blai/projs/soupscript/samples.htm"
        Case 4
            OpenWebpage "http://www.kgv.net/blai/projs/soupscript/commands.htm"
    End Select
End Sub

Private Sub titAboutProg_Click()
    On Error Resume Next
    ShellAbout Me.hWnd, App.CompanyName & "(R) " & App.ProductName & " " & MyVer & " (V." & App.Major & "." & App.Minor & "." & App.Revision & ")", App.LegalCopyright & vbCrLf & "Written by Brian Lai", Me.Icon
End Sub

Public Sub titAboutUpdateNow_Click()
    'On Error Resume Next
    Update
End Sub

Private Sub titAboutUpdatesStartup_Click()
    On Error Resume Next
    SaveSet "StartupUpdate", IIf(titAboutUpdatesStartup.Checked, "0", "1")
    RefreshMenuCheckBoxes
End Sub

Private Sub titAboutWeb_Click()
    On Error Resume Next
    OpenWebpage "http://www.kgv.net/blai/projs/soupscript"
End Sub

Private Sub titDebugFlush_Click()
    On Error Resume Next
    If MsgBox("Are you sure you want to delete the program DLL?" & vbCrLf & vbCrLf & _
                    "If you do, it will be created again when you build EXEs.", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    Kill FindPath(App.Path, "SoupBase.dll")
End Sub

Private Sub titDebugNoteArray_Click(Index As Integer)
    On Error Resume Next
    SaveSet "Verbose", CStr(Index)
    DebugLevel = Index
    RefreshMenuCheckBoxes
End Sub

Private Sub titDebugRuntimeArr_Click(Index As Integer)
    On Error Resume Next
    SaveSet "MakeDLL", CStr(Index)
    RefreshMenuCheckBoxes
End Sub

Public Sub titDebugStop_Click()
    KillEXE
End Sub

Private Sub titDebugStrict_Click()
    On Error Resume Next
    SaveSet "Strict", IIf(titDebugStrict.Checked, "0", "1")
    RefreshMenuCheckBoxes
End Sub

Private Sub titEditArray_Click(Index As Integer)
    On Error Resume Next
    With AF.txtCode
        Select Case Index
            Case 3
                .SelStart = 0
                .SelLength = Len(.Text)
            Case 0
                Clipboard.SetText .SelText
                .SelText = ""
            Case 1
                Clipboard.SetText .SelText
            Case 2
                .SelText = Clipboard.GetText
        End Select
    End With
End Sub

Private Sub titEditCodeBar_Click()
    On Error Resume Next
    picCMDPanel.Width = IIf(picCMDPanel.Width = Bar0, Bar0Width, Bar0)
End Sub

Public Sub titFileBuild_Click()
    'On Error Resume Next
    With cmndlg
        .filefilter = "Applications (*.exe)|*.exe"
        .flags = 5 Or 2
        SaveFile
        If Len(.FileName) = 0 Then Exit Sub
        If LCase$(Right$(.FileName, 4)) <> ".exe" Then .FileName = .FileName & ".exe" 'autoname
        'AF.Caption = .FileName
        Compiler "", .FileName
    End With
End Sub

Private Sub titFileClose_Click()
    On Error Resume Next
    Unload AF
End Sub

Private Sub titFileDeploy_Click()
    On Error Resume Next
    If MsgBox("Deployment reduces the size and readability of your script." & vbCrLf & vbCrLf & "Are you sure you want to deploy this script?", vbYesNo + vbQuestion) = vbYes Then
'        AF.txtCode.Text = Replace(AF.txtCode.Text, vbCrLf, "")
        AF.txtCode.Text = Replace(AF.txtCode.Text, vbTab & vbTab, vbTab)
        AF.txtCode.Text = Replace(AF.txtCode.Text, vbTab & vbTab, vbTab)
        AF.txtCode.Text = Replace(AF.txtCode.Text, vbTab, " ")
        AF.txtCode.Text = CodeFriendly(AF.txtCode.Text)
    End If
End Sub

Public Sub titFileNew_Click()
    On Error Resume Next
    Dim K As Form
    Set K = New frmScript
    K.Show
End Sub

Public Sub titFileOpen_Click()
    On Error Resume Next
    With cmndlg
        .filefilter = "Script File (*.script)|*.script|All Files (*.*)|*.*"
        OpenFile
        If Len(.FileName) = 0 Then Exit Sub
        
        If Dir(.FileName) <> "" Then LoadFile .FileName
    End With
End Sub

Private Sub titFileSave_Click()
    On Error Resume Next
    If AF.Caption = "*" Then '1 is the *
        titFileSaveAs_Click 'if theres no name... save as!
        Exit Sub
    End If
    If Right$(AF.Caption, 1) = "*" Then AF.Caption = Left$(AF.Caption, Len(AF.Caption) - 1)
    TXTFileSave AF.txtCode.Text, AF.Caption
End Sub

Private Sub titFileSaveAs_Click()
    With cmndlg
        .filefilter = "script file (*.script)|*.script"
        .flags = 5 Or 2
        SaveFile
        If Len(.FileName) = 0 Then Exit Sub
        If LCase$(Right$(.FileName, 7)) <> ".script" Then .FileName = .FileName & ".script" 'autoname
        AF.Caption = .FileName
    End With
    titFileSave_Click
End Sub

Function SetFileName()
    On Error Resume Next
    With cmndlg
        .filefilter = "script file (*.script)|*.script"
        .flags = 5 Or 2
        SaveFile
        If Len(.FileName) = 0 Then Exit Function
        AF.Caption = .FileName
    End With
End Function

Function LoadFile(FileN As String)
    On Error Resume Next
    Dim K As Form
    Dim FF As Integer
    Dim tmp As String, REAL As String
    
    Set K = New frmScript
    FF = FreeFile
    Open FileN For Input As #FF
        Do
            Line Input #FF, tmp
            REAL = REAL & tmp & vbCrLf
        Loop Until EOF(FF)
    Close #FF
    K.txtCode.Text = REAL
    K.Caption = FileN
    K.Show
End Function

Function RunCode(Optional StartLine As Long)
    On Error Resume Next
    Dim newFN As String
    DebugLevel = GetSet("Verbose", "1")
    SStatus App.ProductName & ": Initializing parser", 0
    
    PutDLL
    
    newFN = Compiler
    If Len(newFN) > 0 Then
        ShellExecute Me.hWnd, "", newFN, "", "", 0 ', vbNormalFocus
        SStatus "Program run", 0
    'Else
    '    SStatus App.ProductName & " will only compile if you have code.", 2
    End If
    
    'SStatus
End Function

Function Compiler(Optional ByVal TheCode As String, Optional ByVal newFN As String, Optional ByVal Silent As Boolean = False) As String
    On Local Error GoTo errTrap
    
    Dim FF As Integer, FF2 As Integer
    Dim Time1 As Long, Time2 As Long ', V2 As Long
    Dim myDLL As String
    
    If GetDRM("DisallowCompile") = "1" Then
        If GetSet("DisallowCompile", "1") = "1" Then
            MsgBox App.ProductName & " cannot build this program due to regional violation.", vbCritical
            Exit Function
        End If
    End If

    
    Time1 = GetTickCount
    
    If Len(TheCode) = 0 Then TheCode = AF.txtCode.Text
    TheCode = CodeFriendly(TheCode)
    If Len(TheCode) = 0 Then Exit Function
    TheCode = "[SOUP]" & TheCode
    
    myDLL = FindPath(App.Path, StdFN)
    If Len(newFN) = 0 Then newFN = FindPath(GetTempDir, "SoupBase.exe")
    
    
    FileCopy myDLL, newFN    'first copy that file to the user provided file name.
    
    FF = FreeFile
        Open newFN For Binary As #FF        'Open DLL File
        Put #FF, LOF(FF), TheCode
    Close #FF                            'Close Application
    
    Time2 = GetTickCount
    
    Compiler = newFN 'if the build succeeds, this function returns the path
    
    If Not Silent Then SStatus "Program built (" & Round((Time2 - Time1) / 1000, 2) & "s)", 0
errTrap:
    If Err.Number <> 0 And Not Silent Then
        SStatus "Build error " & Err.Number & ": " & Err.Description, 2
    End If
End Function

Private Sub titFileWizard_Click()
    On Error Resume Next
    frmWizard.Hello = True
    frmWizard.Show 1
End Sub

Public Sub titRunArray_Click(Index As Integer)
    On Error Resume Next
    AF.txtCode.Locked = True
    Select Case Index
        Case 0
            RunCode
        Case 1
            RunCode GetLinePos(AF.txtCode)
    End Select
    AF.txtCode.Locked = False
End Sub

Private Sub titViewClrArray_Click(Index As Integer)
    'On Error Resume Next
    Dim myClrs(5) As Long
    Dim I As Long
    
    Select Case Index
        Case 0 'white
            myClrs(0) = vbWhite 'backcolor0
            myClrs(1) = vbWhite 'backcolor1
            myClrs(2) = vbWhite 'backcolor2
            myClrs(3) = vbBlack 'forecolor0
            myClrs(4) = vbBlack 'forecolor1
            myClrs(5) = vbBlack 'forecolor2
        Case 1 'black
            myClrs(0) = RGB(40, 40, 40) 'backcolor0
            myClrs(1) = RGB(40, 40, 40) 'backcolor1
            myClrs(2) = RGB(40, 40, 40) 'backcolor2
            myClrs(3) = vbWhite 'forecolor0
            myClrs(4) = RGB(150, 255, 150) 'forecolor1
            myClrs(5) = vbWhite 'forecolor2
        Case 2 'dark orange
            myClrs(0) = RGB(97, 50, 0) 'backcolor0
            myClrs(1) = RGB(97, 50, 0) 'backcolor1
            myClrs(2) = RGB(97, 50, 0) 'backcolor2
            myClrs(3) = vbWhite 'forecolor0
            myClrs(4) = RGB(255, 255, 220) 'forecolor1
            myClrs(5) = vbWhite 'forecolor2
        Case 3 'white
            myClrs(0) = RGB(30, 40, 77) 'backcolor0
            myClrs(1) = vbButtonFace  'backcolor1
            myClrs(2) = RGB(30, 40, 77) 'backcolor2
            myClrs(3) = vbWhite 'forecolor0
            myClrs(4) = vbButtonText  'forecolor1
            myClrs(5) = vbWhite 'forecolor2
    End Select
    
    For I = 0 To 5 Step 1
        If I <= 2 Then
            SaveSet "Backcolor" & I, CStr(myClrs(I))
        ElseIf I > 2 And I <= 5 Then
            SaveSet "Forecolor" & I - 3, CStr(myClrs(I))
        End If
    Next
    ColorizeMe
    'MsgBox "Restart to take effect.", vbInformation
End Sub

Private Sub titViewDebug_Click()
    On Error Resume Next
    With AF.pDP
        .Height = IIf(.Height <> AF.ScaleHeight / 2, AF.ScaleHeight / 2, AF.lblDrag.Height)
    End With
End Sub

Private Sub titViewFont_Click()
    On Error Resume Next
    Dim K As String
        
    SelectFont.mFontName = AF.txtCode.FontName
    
    ShowFont
    
    K = SelectFont.mFontName
    If Len(K) > 0 And AF.txtCode.FontName <> K Then
        SaveSet "ScriptFont", K
        SaveSet "CodeFont", K
        SaveSet "DebugFont", K
        MsgBox "Restart to take effect.", vbInformation
    End If
End Sub

Function RefreshMenuCheckBoxes()
    On Error Resume Next
    Dim I As Integer
    Dim Index As Integer
    
    Index = Val(GetSet("MakeDLL", "0"))
    For I = titDebugRuntimeArr.LBound To titDebugRuntimeArr.UBound Step 1
        titDebugRuntimeArr(I).Checked = (I = Index)
    Next
    
    Index = Val(GetSet("Verbose", "1"))
    For I = titDebugNoteArray.LBound To titDebugNoteArray.UBound Step 1
        titDebugNoteArray(I).Checked = (I = Index)
    Next
    
    Index = Val(GetSet("StartupUpdate", "1"))
    titAboutUpdatesStartup.Checked = (Index = 1)
    
    Index = Val(GetSet("Strict", "0"))
    titDebugStrict.Checked = (Index = 1)
    
End Function
