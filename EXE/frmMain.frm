VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "SoupBase"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6885
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   6885
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Visible         =   0   'False
   Begin VB.TextBox TXTImmediate 
      BackColor       =   &H00000000&
      BorderStyle     =   0  '¨S¦³®Ø½u
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   4215
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  '¨âªÌ¬Ò¦³
      TabIndex        =   0
      Text            =   "frmMain.frx":058A
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    On Error Resume Next
    Me.Caption = App.ProductName
    TXTImmediate.Text = Replace(TXTImmediate.Text, "%ver%", MyVer)
    SDebug
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    TXTImmediate.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End 'i dont want you anymore, really
End Sub

Private Sub TXTImmediate_KeyPress(KeyAscii As Integer)
    AnyKey = KeyAscii
End Sub
