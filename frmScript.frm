VERSION 5.00
Begin VB.Form frmScript 
   Caption         =   "*"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmScript.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   6840
   WindowState     =   2  '³Ì¤j¤Æ
   Begin VB.PictureBox picAC 
      Appearance      =   0  '¥­­±
      BackColor       =   &H80000018&
      BorderStyle     =   0  '¨S¦³®Ø½u
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   960
      ScaleHeight     =   1095
      ScaleWidth      =   3255
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   3255
      Begin VB.ListBox lstAC 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   -15
         TabIndex        =   2
         Top             =   -15
         Width           =   3270
      End
      Begin VB.Label lblLengthGetter 
         AutoSize        =   -1  'True
         Caption         =   "blah"
         Height          =   195
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   360
      End
      Begin VB.Label lblCodeHint 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Code Hint"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   0
         TabIndex        =   3
         Top             =   840
         Width           =   3225
         WordWrap        =   -1  'True
      End
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   4575
      Left            =   0
      MultiLine       =   -1  'True
      OLEDropMode     =   1  '¤â°Ê
      ScrollBars      =   2  '««ª½±²¶b
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'[frmScript]
Option Explicit

Dim MyY As Long

Private Sub Form_Activate()
    On Error Resume Next
    InitCommonControls
End Sub

Private Sub Form_Load()
    On Error Resume Next
    With txtCode
        .BackColor = CLng(GetSet("BackColor0", CStr(RGB(40, 40, 40))))
        .ForeColor = CLng(GetSet("ForeColor0", CStr(RGB(255, 255, 255))))
        .FontName = GetSet("ScriptFont", "Verdana")
        .FontSize = CLng(GetSet("ScriptFontSize", "10"))
    End With
    SkinFormEx Me
End Sub

Private Sub Form_Paint()
    On Error Resume Next
    Me.Tag = MyFormNo
    MyFormNo = MyFormNo + 1
    SyncTabs False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Dim nFile As String
    Me.SetFocus
    If Right$(Me.Caption, 1) = "*" Then 'unsaved
        Select Case MsgBox("Unsaved document. Save before closing?", vbYesNoCancel + vbQuestion)
            Case vbYes
                If Me.Caption = "*" Then '1 is the *
                    With cmndlg
                        .filefilter = "script file (*.script)|*.script"
                        .flags = 5 Or 2
                        SaveFile
                        If Len(.FileName) = 0 Then Exit Sub
                        If LCase$(Right$(.FileName, 7)) <> ".script" Then .FileName = .FileName & ".script" 'autoname
                        Me.Caption = .FileName
                    End With
                    If Right$(Me.Caption, 1) = "*" Then Me.Caption = Left$(Me.Caption, Len(Me.Caption) - 1)
                    TXTFileSave Me.txtCode.Text, Me.Caption
                End If
            Case vbNo
                'do nothing
            Case vbCancel
                Cancel = 1
        End Select
    End If
End Sub

Public Sub Form_Resize()
    On Error Resume Next
    txtCode.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight '- pDP.Height
End Sub

Private Sub Form_Terminate()
    UnloadApp
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Tag = "CLOSED"
    SyncTabs
End Sub

Private Sub lstAC_Click()
    On Error Resume Next
    Dim K As String, J As String
    K = lstAC.List(lstAC.ListIndex)
    If InStr(1, K, "(") > 0 Then
        K = Left$(lstAC.List(lstAC.ListIndex), InStr(1, lstAC.List(lstAC.ListIndex), "(") - 1)
    Else
        K = Left$(K, Len(K) - 1) 'removing the ;
    End If
    J = GetHint(K, True)
    If Len(J) > 0 Then
        lblCodeHint.Caption = K & vbCrLf & Replace(J, "\n", vbCrLf)
        lblCodeHint.Visible = True
    Else
        lblCodeHint.Visible = False
    End If
End Sub

Private Sub lstAC_DblClick()
    On Error Resume Next
    Dim ThisLine As String
    Dim iiA As Long
    
    ThisLine = LCase$(GetTextboxLine(txtCode, GetLinePos(txtCode)))
    iiA = InStr(1, ThisLine, "=")
    
    With lstAC
        txtCode.SelText = Right$(.List(.ListIndex), Len(.List(.ListIndex)) - Len(ThisLine) + iiA)  'add the remainder
    End With
    picAC.Visible = False
End Sub

'Private Sub lstAC_DblClick()
'    On Error Resume Next
'    Dim lngOldCharPos As Long
'    Dim strLineStart As String
'    Dim iiA As Long, iiB As Long, iiC As Long
'    lngOldCharPos = txtCode.SelStart
'    strLineStart = GetTextboxLine(txtCode, GetLinePos(txtCode))
'
'    iiA = InStr(1, strLineStart, "=")
''    iiB = InStr(1, strLineStart, " ")
''    iiC = InStr(1, strLineStart, "(")
'
'    If iiA = 0 Then iiA = 1 Else iiA = iiA + 1
''    If iiC <> 0 Then 'if bracket exists
''        If iiC >= iiB Then 'and after space
''            iiC = iiB
''        End If
''    End If
''    If iiA > iiC Then iiA = iiC
'
'    strLineStart = Mid$(strLineStart, iiA)
'    With lstAC
'        txtCode.SelText = StripLineDescriptors(Mid$(.List(.ListIndex), Len(strLineStart) + 1))
'        picAC.Visible = False
'    End With
'End Sub

'
Private Sub txtCode_Change()
    On Error Resume Next
    If Right$(Me.Caption, 1) <> "*" Then Me.Caption = Me.Caption & "*"
'    txtCode_Click
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Dim K() As String
    Dim ThisLine As String
    Dim I As Long, J As Long, L As Long
    Dim IntelliStr As String ', IS22 As String
    
    ThisLine = GetTextboxLine(txtCode, GetLinePos(txtCode))
    SStatus "Line " & GetLinePos(txtCode), 0

    
    Select Case KeyAscii
        Case vbKeyTab
            txtCode.SelText = vbTab
            KeyAscii = 0
    End Select
End Sub

Function AutoCompleteCMDs(FirstChars As String)
    On Error Resume Next
    Dim K As String
    Dim I As Long
    Dim eX As Long, eY As Long
    Dim MaxL As Long
    
    
    MaxL = 3600 'minimum width of tooltip
    If Len(FirstChars) = 0 Then
        picAC.Visible = False
    Else
        lstAC.Clear
        With frmMain.lstCMDs
            For I = 0 To .ListCount - 1 Step 1
                If LCase$(Left$(.List(I), Len(FirstChars))) = LCase$(FirstChars) Then
                    K = StripLineDescriptors(.List(I))
                    lstAC.AddItem K
                    lblLengthGetter.Caption = K
                    If lblLengthGetter.Width > MaxL Then MaxL = lblLengthGetter.Width
                End If
            Next
        End With
        
        With lstAC
            If .ListCount > 0 Then
                'prevents "tooltip" from moving out of screen
                eX = CaretX(txtCode)
                If eX + lstAC.Width > Me.ScaleWidth Then eX = Me.ScaleWidth - lstAC.Width
                eY = CaretY(txtCode)
                If eY + lstAC.Height > Me.ScaleHeight Then eY = Me.ScaleHeight - lstAC.Height
                'end move prevention
                
                .ListIndex = 0
                lstAC_Click 'call the memory
                picAC.Move eX, eY + 300, MaxL, lblCodeHint.Top + IIf(lblCodeHint.Visible, lblCodeHint.Height, -45) + 30
                .Width = picAC.Width + 30
                lblCodeHint.Width = .Width

                picAC.Visible = True
            Else
                picAC.Visible = False
            End If
        End With
    End If
End Function

Private Sub txtCode_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim iiA As Long, iiB As Long, iiC As Long
    Dim ThisLine As String
    
    ThisLine = GetTextboxLine(txtCode, GetLinePos(txtCode))
    iiA = InStr(1, ThisLine, "=")
    If iiA = 0 Then 'if it's found
        iiA = 1
    End If
    iiB = InStr(iiA, ThisLine, "(")
    iiC = InStr(iiA, ThisLine, " ")
    
    If iiC > iiB And iiB <> 0 Then iiC = iiB
    If iiC = 0 Then iiC = Len(ThisLine) + 1
    ThisLine = Mid$(ThisLine, iiA, iiC - iiA)
    ThisLine = Replace(ThisLine, "=", "")
    
    AutoCompleteCMDs ThisLine
    
'    txtCode_Click
    
    
    'Alt+> autocomplete
    Dim I As Long
    Dim strLineStart As Long
    Dim K As String
    
    If Shift = 4 Then
        If KeyCode = vbKeyRight Then
        
            strLineStart = GetTextboxLine(txtCode, GetLinePos(txtCode))
            iiA = InStr(1, strLineStart, "=")
            If iiA = 0 Then iiA = 1 Else iiA = iiA + 1
            
            With frmMain.lstCMDs
                For I = 0 To .ListCount - 1 Step 1
                    K = ThisLine
                    If LCase$(Left$(.List(I), Len(K))) = K Then
                        txtCode.SelText = StripLineDescriptors(Right$(.List(I), Len(.List(I)) - Len(K))) 'add the remainder
                        Exit For
                    End If
                Next
            End With
        End If
    End If
End Sub

Private Sub txtCode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    AutoCompleteCMDs Trim$(txtCode.SelText)
End Sub

Private Sub txtCode_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    txtCode.SelText = Data.GetData(1)
End Sub
