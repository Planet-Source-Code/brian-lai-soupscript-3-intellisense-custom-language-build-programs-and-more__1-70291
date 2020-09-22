VERSION 5.00
Begin VB.Form frmWizard 
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "Wizard"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   Begin VB.CheckBox chkShowMe 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show this everytime"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4800
      Width           =   2775
   End
   Begin VB.CommandButton btnAction 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   2
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton btnAction 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   1
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Frame frmFileProp 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Caption         =   "New file"
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   4335
         Left            =   120
         ScaleHeight     =   4335
         ScaleWidth      =   5415
         TabIndex        =   3
         Top             =   120
         Width           =   5415
         Begin VB.OptionButton Opts 
            Caption         =   "Create blank project"
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   6
            Top             =   360
            Value           =   -1  'True
            Width           =   4095
         End
         Begin VB.OptionButton Opts 
            Caption         =   "Create sample ""hello world"" script"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   5
            Top             =   720
            Width           =   4095
         End
         Begin VB.OptionButton Opts 
            Caption         =   "Open an existing script"
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   4
            Top             =   1680
            Width           =   4095
         End
         Begin VB.Frame Frm 
            Caption         =   "New file"
            Height          =   1215
            Index           =   0
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   5415
            Begin VB.Image Image1 
               Height          =   240
               Index           =   0
               Left            =   360
               Picture         =   "frmWizard.frx":000C
               Top             =   480
               Width           =   240
            End
         End
         Begin VB.Frame Frm 
            Caption         =   "Open file"
            Height          =   3015
            Index           =   1
            Left            =   0
            TabIndex        =   8
            Top             =   1320
            Width           =   5415
            Begin VB.FileListBox fileChoose 
               Enabled         =   0   'False
               Height          =   2190
               Left            =   840
               Pattern         =   "*.script"
               TabIndex        =   9
               Top             =   720
               Width           =   4455
            End
            Begin VB.Image Image1 
               Height          =   240
               Index           =   1
               Left            =   360
               Picture         =   "frmWizard.frx":03CB
               Top             =   360
               Width           =   240
            End
         End
      End
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '¤£³z©ú
      BorderStyle     =   0  '³z©ú
      Height          =   615
      Left            =   0
      Top             =   4560
      Width           =   5655
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Hello As Boolean

Private Sub btnAction_Click(Index As Integer)
    On Error Resume Next
    Dim I As Long, J As Long
    Me.Hide
    If Index = 0 Then
    For I = 0 To Opts.UBound Step 1
        J = J - Opts(I).Value * I
    Next
    Debug.Print J
    Select Case J
        Case 0
            frmMain.titFileNew_Click
        Case 1
            frmMain.titFileNew_Click
            AF.txtCode.Text = "CLS;REM this line clears the screen, and REM is a comment;" & vbCrLf & "COut(Hello World!);REM this line shows the message;" & vbCrLf & "Break;REM this line pauses the code so you can see the message;"
            
        Case 2
            frmMain.LoadFile FindPath(fileChoose.Path, fileChoose.FileName)
        Case Else
            MsgBox "How the heck did you choose that?!", vbCritical
    End Select
    End If
    SaveSet "StartupWizard", CStr(chkShowMe.Value)
    Unload Me
End Sub

Private Sub fileChoose_DblClick()
    On Error Resume Next
    btnAction_Click 0
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    InitCommonControls
End Sub

Private Sub Form_Load()
    On Error Resume Next
    chkShowMe.Value = Val(GetSet("StartupWizard", "1"))
    If chkShowMe.Value <> 0 Or Hello = True Then
        SkinFormEx Me
        fileChoose.Path = App.Path 'if it isn't already
    Else
        Unload Me
    End If
End Sub

Private Sub Opts_Click(Index As Integer)
    fileChoose.Enabled = (Index = 2)
End Sub

Private Sub Form_Terminate()
    On Error Resume Next
    UnloadApp
End Sub

Private Sub Opts_DblClick(Index As Integer)
    fileChoose_DblClick
End Sub
