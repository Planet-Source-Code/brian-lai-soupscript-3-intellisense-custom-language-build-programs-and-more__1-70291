VERSION 5.00
Begin VB.UserControl CB 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2205
   DefaultCancel   =   -1  'True
   OLEDropMode     =   1  '¤â°Ê
   PropertyPages   =   "CB.ctx":0000
   ScaleHeight     =   107
   ScaleMode       =   3  '¹³¯À
   ScaleWidth      =   147
   ToolboxBitmap   =   "CB.ctx":004A
   Begin VB.Timer OverTimer 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "CB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const isOCX = False

Private Const cbVersion As String = "2.0.6"

'CHAMELEON BUTTON copyright     2001-2002 by gonchuki      E -mail: gonchuki@ yahoo.es
'This is not the normal Cham button - Brian Lai chose to edit it and get everything he thought was useless out of the control.

'Removed functions:
'many styles, focus rect function, special effects, colour schemes, soft bevel, mask colour toggle, greyscale, hand pointer

'Edited functions:
'mouseover leaking problem. now it doesn't (except when in Windows 32-bit button mode)

Private Declare Function SetPixel Lib "gdi32" Alias "SetPixelV" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNDKSHADOW = 21
Private Const COLOR_BTNLIGHT = 22

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10
Private Const DT_CENTER = &H1 Or DT_WORDBREAK Or &H4

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Const PS_SOLID = 0

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Const RGN_DIFF = 4

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBTRIPLE
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBTRIPLE
End Type

Public Enum ButtonTypes
    [Windows 32-bit] = 2    'the classic windows button
    [Java metal] = 5        'there are also other styles but not so different from windows one
    [Simple Flat] = 7       'the standard flat button seen on toolbars
    [Flat Highlight] = 8    'again the flat button but this one has no border until the mouse is over it
    [Office XP] = 9         'the new Office XP button
End Enum

Public Enum ColorTypes
    [Use Windows] = 1
    [Custom] = 2
End Enum

Public Enum PicPositions
    cbLeft = 0
    cbRight = 1
    cbTop = 2
    cbBottom = 3
    cbBackground = 4
End Enum

'events
Public Event Click()
Attribute Click.VB_MemberFlags = "200"
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseOver()
Public Event MouseOut()
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

'variables
Private MyButtonType As ButtonTypes
Private MyColorType As ColorTypes
Private PicPosition As PicPositions
Private He As Long  'the height of the button
Private Wi As Long  'the width of the button
Private BackC As Long 'back color
Private BackO As Long 'back color when mouse is over
Private ForeC As Long 'fore color
Private ForeO As Long 'fore color when mouse is over
Private MaskC As Long 'mask color
Private OXPb As Long, OXPf As Long
Private picNormal As StdPicture, picHover As StdPicture
Private pDC As Long, pBM As Long, oBM As Long 'used for the treansparent button
Private elTex As String     'current text
Private rc As RECT, rc2 As RECT, rc3 As RECT, fc As POINTAPI 'text and focus rect locations
Private picPT As POINTAPI, picSZ As POINTAPI  'picture Position & Size
Private rgnNorm As Long
Private LastButton As Byte, LastKeyDown As Byte
Private isEnabled As Boolean
Private HasFocus As Boolean
Private cFace As Long, cLight As Long, cHighLight As Long, cShadow As Long, cDarkShadow As Long, cText As Long, cTextO As Long, cFaceO As Long, cMask As Long, XPFace As Long
Private lastStat As Byte, TE As String, isShown As Boolean  'used to avoid unnecessary repaints
Private isOver As Boolean, inLoop As Boolean
Private captOpt As Long
Private isCheckbox As Boolean, cValue As Boolean

Private Sub OverTimer_Timer()
    On Error Resume Next
    If Not isMouseOver Then
        OverTimer.Enabled = False
        isOver = False
        Call Redraw(0, True)
        RaiseEvent MouseOut
    End If
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    On Error Resume Next
    LastButton = 1
    Call UserControl_Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    On Error Resume Next
    If Not MyColorType = [Custom] Then
        Call SetColors
        Call Redraw(lastStat, True)
    End If
End Sub

Private Sub UserControl_Click()
    On Error Resume Next
    If LastButton = 1 And isEnabled Then
        If isCheckbox Then cValue = Not cValue
        Call Redraw(0, True) 'be sure that the normal status is drawn
        UserControl.Refresh
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_DblClick()
    On Error Resume Next
    If LastButton = 1 Then
        Call UserControl_MouseDown(1, 0, 0, 0)
        SetCapture hWnd
    End If
End Sub

Private Sub UserControl_GotFocus()
    On Error Resume Next
    HasFocus = True
    Call Redraw(lastStat, True)
End Sub

Private Sub UserControl_Hide()
    On Error Resume Next
    isShown = False
End Sub

Private Sub UserControl_Initialize()
    On Error Resume Next
    'this makes the control to be slow, remark this line if the "not redrawing" problem is not important for you: ie, you intercept the Load_Event (with breakpoint or messageBox) and the button does not repaint...
    isShown = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    RaiseEvent KeyDown(KeyCode, Shift)
    
    LastKeyDown = KeyCode
    Select Case KeyCode
        Case 32 'spacebar pressed
            Call Redraw(2, False)
        Case 39, 40 'right and down arrows
            SendKeys "{Tab}"
        Case 37, 38 'left and up arrows
            SendKeys "+{Tab}"
    End Select
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)

If (KeyCode = 32) And (LastKeyDown = 32) Then 'spacebar pressed, and not cancelled by the user
    On Error Resume Next
    If isCheckbox Then cValue = Not cValue
    Call Redraw(0, False)
    UserControl.Refresh
    RaiseEvent Click
End If
End Sub

Private Sub UserControl_LostFocus()
    On Error Resume Next
    HasFocus = False
    Call Redraw(lastStat, True)
End Sub

Private Sub UserControl_InitProperties()
    On Error Resume Next
    isEnabled = True
    elTex = Ambient.DisplayName
    Set UserControl.Font = Ambient.Font
    MyButtonType = [Windows 32-bit]
    MyColorType = [Use Windows]
    Call SetColors
    BackC = cFace: BackO = BackC
    ForeC = cText: ForeO = ForeC
    MaskC = &HC0C0C0
    Call CalcTextRects
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    RaiseEvent MouseDown(Button, Shift, X, Y)
    LastButton = Button
    If Button <> 2 Then Call Redraw(2, False)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Button < 2 Then
        If Not isMouseOver Then
            'we are outside the button
            Call Redraw(0, False)
        Else
            'we are inside the button
            If Button = 0 And Not isOver Then
                OverTimer.Enabled = True
                isOver = True
                Call Redraw(0, True)
                RaiseEvent MouseOver
            ElseIf Button = 1 Then
                isOver = True
                Call Redraw(2, False)
                isOver = False
            End If
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Button <> 2 Then Call Redraw(0, False)
End Sub

'########## BUTTON PROPERTIES ##########
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
    On Error Resume Next
    BackColor = BackC
End Property
Public Property Let BackColor(ByVal theCol As OLE_COLOR)
    On Error Resume Next
    BackC = theCol
    If Not Ambient.UserMode Then BackO = theCol
    Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "BCOL"
End Property

Public Property Get BackOver() As OLE_COLOR
Attribute BackOver.VB_ProcData.VB_Invoke_Property = ";Appearance"
    On Error Resume Next
    BackOver = BackO
End Property
Public Property Let BackOver(ByVal theCol As OLE_COLOR)
    On Error Resume Next
    BackO = theCol
    Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "BCOLO"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
    On Error Resume Next
    ForeColor = ForeC
End Property
Public Property Let ForeColor(ByVal theCol As OLE_COLOR)
    On Error Resume Next
    ForeC = theCol
    If Not Ambient.UserMode Then ForeO = theCol
    Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "FCOL"
End Property

Public Property Get ForeOver() As OLE_COLOR
Attribute ForeOver.VB_ProcData.VB_Invoke_Property = ";Appearance"
    On Error Resume Next
    ForeOver = ForeO
End Property
Public Property Let ForeOver(ByVal theCol As OLE_COLOR)
    On Error Resume Next
    ForeO = theCol
    Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "FCOLO"
End Property

Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    On Error Resume Next
    MaskColor = MaskC
End Property
Public Property Let MaskColor(ByVal theCol As OLE_COLOR)
    On Error Resume Next
    MaskC = theCol
    Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "MCOL"
End Property

Public Property Get ButtonType() As ButtonTypes
Attribute ButtonType.VB_ProcData.VB_Invoke_Property = ";Appearance"
    On Error Resume Next
    ButtonType = MyButtonType
End Property

Public Property Let ButtonType(ByVal newValue As ButtonTypes)
    On Error Resume Next
    MyButtonType = newValue
    Call UserControl_Resize
    PropertyChanged "BTYPE"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Caption.VB_UserMemId = 0
    On Error Resume Next
    Caption = elTex
End Property

Public Property Let Caption(ByVal newValue As String)
    On Error Resume Next
    elTex = newValue
    Call SetAccessKeys
    Call CalcTextRects
    Call Redraw(0, True)
    PropertyChanged "TX"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    On Error Resume Next
    Enabled = isEnabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    On Error Resume Next
    isEnabled = newValue
    Call Redraw(0, True)
    UserControl.Enabled = isEnabled
    PropertyChanged "ENAB"
End Property

Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    On Error Resume Next
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByRef newFont As Font)
    On Error Resume Next
    Set UserControl.Font = newFont
    Call CalcTextRects
    Call Redraw(0, True)
    PropertyChanged "FONT"
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_MemberFlags = "400"
    On Error Resume Next
    FontBold = UserControl.FontBold
End Property

Public Property Let FontBold(ByVal newValue As Boolean)
    On Error Resume Next
    UserControl.FontBold = newValue
    Call CalcTextRects
    Call Redraw(0, True)
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_MemberFlags = "400"
    On Error Resume Next
    FontItalic = UserControl.FontItalic
End Property

Public Property Let FontItalic(ByVal newValue As Boolean)
    On Error Resume Next
    UserControl.FontItalic = newValue
    Call CalcTextRects
    Call Redraw(0, True)
End Property

Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_MemberFlags = "400"
    On Error Resume Next
    FontUnderline = UserControl.FontUnderline
End Property

Public Property Let FontUnderline(ByVal newValue As Boolean)
    On Error Resume Next
    UserControl.FontUnderline = newValue
    Call CalcTextRects
    Call Redraw(0, True)
End Property

Public Property Get FontSize() As Integer
Attribute FontSize.VB_MemberFlags = "400"
    On Error Resume Next
    FontSize = UserControl.FontSize
End Property

Public Property Let FontSize(ByVal newValue As Integer)
    On Error Resume Next
    UserControl.FontSize = newValue
    Call CalcTextRects
    Call Redraw(0, True)
End Property

Public Property Get FontName() As String
Attribute FontName.VB_MemberFlags = "400"
    On Error Resume Next
    FontName = UserControl.FontName
End Property

Public Property Let FontName(ByVal newValue As String)
    On Error Resume Next
    UserControl.FontName = newValue
    Call CalcTextRects
    Call Redraw(0, True)
End Property

Public Property Get ColorScheme() As ColorTypes
Attribute ColorScheme.VB_ProcData.VB_Invoke_Property = ";Appearance"
    On Error Resume Next
    ColorScheme = MyColorType
End Property

Public Property Let ColorScheme(ByVal newValue As ColorTypes)
    On Error Resume Next
    MyColorType = newValue
    Call SetColors
    Call Redraw(0, True)
    PropertyChanged "COLTYPE"
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Appearance"
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal newPointer As MousePointerConstants)
    On Error Resume Next
    UserControl.MousePointer = newPointer
    PropertyChanged "MPTR"
End Property

Public Property Get MouseIcon() As StdPicture
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = ";Appearance"
    On Error Resume Next
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal newIcon As StdPicture)
    On Local Error Resume Next
    Set UserControl.MouseIcon = newIcon
    PropertyChanged "MICON"
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    On Error Resume Next
    hWnd = UserControl.hWnd
End Property

Public Property Get PictureNormal() As StdPicture
Attribute PictureNormal.VB_ProcData.VB_Invoke_Property = ";Appearance"
    On Error Resume Next
    Set PictureNormal = picNormal
End Property
Public Property Set PictureNormal(ByVal newPic As StdPicture)
    On Error Resume Next
    Set picNormal = newPic
    Call CalcPicSize
    Call CalcTextRects
    Call Redraw(lastStat, True)
    PropertyChanged "PICN"
End Property

Public Property Get PictureOver() As StdPicture
Attribute PictureOver.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set PictureOver = picHover
End Property
Public Property Set PictureOver(ByVal newPic As StdPicture)
    On Error Resume Next
    Set picHover = newPic
    If isOver Then Call Redraw(lastStat, True) 'only redraw i we need to see this picture immediately
    PropertyChanged "PICO"
End Property

Public Property Get PicturePosition() As PicPositions
Attribute PicturePosition.VB_ProcData.VB_Invoke_Property = ";Position"
    On Error Resume Next
    PicturePosition = PicPosition
End Property
Public Property Let PicturePosition(ByVal newPicPos As PicPositions)
    On Error Resume Next
    PicPosition = newPicPos
    PropertyChanged "PICPOS"
    Call CalcTextRects
    Call Redraw(lastStat, True)
End Property

Public Property Get CheckBoxBehaviour() As Boolean
    On Error Resume Next
    CheckBoxBehaviour = isCheckbox
End Property

Public Property Let CheckBoxBehaviour(ByVal newValue As Boolean)
    On Error Resume Next
    isCheckbox = newValue
    Call Redraw(lastStat, True)
    PropertyChanged "CHECK"
End Property

Public Property Get Value() As Boolean
    Value = cValue
End Property

Public Property Let Value(ByVal newValue As Boolean)
    On Error Resume Next
    cValue = newValue
    If isCheckbox Then Call Redraw(0, True)
    PropertyChanged "VALUE"
End Property

Public Property Get Version() As String
    On Error Resume Next
    Version = cbVersion
End Property

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

'########## END OF PROPERTIES ##########

Private Sub UserControl_Resize()
    On Error Resume Next
    If inLoop Then Exit Sub
    'get button size
    GetClientRect UserControl.hWnd, rc3
    'assign these values to He and Wi
    He = rc3.Bottom: Wi = rc3.Right
    'build the FocusRect size and position depending on the button type
    If MyButtonType >= [Simple Flat] And MyButtonType <= 9 Then
        InflateRect rc3, -3, -3
    Else
        InflateRect rc3, -4, -4
    End If
    Call CalcTextRects
    
    If rgnNorm Then DeleteObject rgnNorm
    Call MakeRegion
    SetWindowRgn UserControl.hWnd, rgnNorm, True
    
    If He Then Call Redraw(0, True)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        MyButtonType = .ReadProperty("BTYPE", 2)
        elTex = .ReadProperty("TX", "")
        isEnabled = .ReadProperty("ENAB", True)
        Set UserControl.Font = .ReadProperty("FONT", UserControl.Font)
        MyColorType = .ReadProperty("COLTYPE", 1)
        BackC = .ReadProperty("BCOL", GetSysColor(COLOR_BTNFACE))
        BackO = .ReadProperty("BCOLO", BackC)
        ForeC = .ReadProperty("FCOL", GetSysColor(COLOR_BTNTEXT))
        ForeO = .ReadProperty("FCOLO", ForeC)
        MaskC = .ReadProperty("MCOL", &HC0C0C0)
        UserControl.MousePointer = .ReadProperty("MPTR", 0)
        Set UserControl.MouseIcon = .ReadProperty("MICON", Nothing)
        Set picNormal = .ReadProperty("PICN", Nothing)
        Set picHover = .ReadProperty("PICH", Nothing)
        PicPosition = .ReadProperty("PICPOS", 0)
        isCheckbox = .ReadProperty("CHECK", False)
        cValue = .ReadProperty("VALUE", False)
    End With

    UserControl.Enabled = isEnabled
    Call CalcPicSize
    Call CalcTextRects
    Call SetAccessKeys
End Sub

Private Sub UserControl_Show()
    On Error Resume Next
    isShown = True
    Call SetColors
    Call Redraw(0, True)
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    isShown = False
    DeleteObject rgnNorm
    If pDC Then
        DeleteObject SelectObject(pDC, oBM)
        DeleteDC pDC
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        Call .WriteProperty("BTYPE", MyButtonType)
        Call .WriteProperty("TX", elTex)
        Call .WriteProperty("ENAB", isEnabled)
        Call .WriteProperty("FONT", UserControl.Font)
        Call .WriteProperty("COLTYPE", MyColorType)
        Call .WriteProperty("BCOL", BackC)
        Call .WriteProperty("BCOLO", BackO)
        Call .WriteProperty("FCOL", ForeC)
        Call .WriteProperty("FCOLO", ForeO)
        Call .WriteProperty("MCOL", MaskC)
        Call .WriteProperty("MPTR", UserControl.MousePointer)
        Call .WriteProperty("MICON", UserControl.MouseIcon)
        Call .WriteProperty("PICN", picNormal)
        Call .WriteProperty("PICH", picHover)
        Call .WriteProperty("PICPOS", PicPosition)
        Call .WriteProperty("CHECK", isCheckbox)
        Call .WriteProperty("VALUE", cValue)
    End With
End Sub

Private Sub Redraw(ByVal curStat As Byte, ByVal Force As Boolean)
    On Error Resume Next
    'here is the CORE of the button, everything is drawn here
    'it's not well commented but i think that everything is
    'pretty self explanatory...
        
    If isCheckbox And cValue Then curStat = 2
    
    If Not Force Then  'check drawing redundancy
        If (curStat = lastStat) And (TE = elTex) Then Exit Sub
    End If
    
    If He = 0 Or Not isShown Then Exit Sub   'we don't want errors
    
    lastStat = curStat
    TE = elTex
    
    Dim I As Long, stepXP1 As Single, XPFace2 As Long, tempCol As Long
    
    With UserControl
    .Cls
    If isOver And MyColorType = Custom Then tempCol = BackC: BackC = BackO: SetColors
    
    DrawRectangle 0, 0, Wi, He, cFace
    
    If isEnabled Then
        If curStat = 0 Then
    '#@#@#@#@#@# BUTTON NORMAL STATE #@#@#@#@#@#
            Select Case MyButtonType
                Case 2 'Windows 32-bit
                    Call DrawCaption(Abs(isOver))
                        DrawFrame cHighLight, cDarkShadow, cLight, cShadow, False
                Case 5 'Java
                    DrawRectangle 1, 1, Wi - 1, He - 1, ShiftColor(cFace, &HC)
                    Call DrawCaption(Abs(isOver))
                    DrawRectangle 1, 1, Wi - 1, He - 1, cHighLight, True
                    DrawRectangle 0, 0, Wi - 1, He - 1, ShiftColor(cShadow, -&H1A), True
                    mSetPixel 1, He - 2, ShiftColor(cShadow, &H1A)
                    mSetPixel Wi - 2, 1, ShiftColor(cShadow, &H1A)
                Case 7, 8 'Flat buttons
                    Call DrawCaption(Abs(isOver))
                    If (MyButtonType = [Simple Flat]) Then
                        DrawFrame cHighLight, cShadow, 0, 0, False, True
                    ElseIf isOver Then
                        If MyButtonType = [Flat Highlight] Then
                            DrawFrame cHighLight, cShadow, 0, 0, False, True
                        Else
                            DrawFrame cHighLight, cDarkShadow, cLight, cShadow, False, False
                        End If
                    End If
                Case 9 'Office XP
                    If isOver Then DrawRectangle 1, 1, Wi, He, OXPf
                    Call DrawCaption(Abs(isOver))
                    If isOver Then DrawRectangle 0, 0, Wi, He, OXPb, True
            End Select
            Call DrawPictures(0)
        ElseIf curStat = 2 Then
    '#@#@#@#@#@# BUTTON IS DOWN #@#@#@#@#@#
            Select Case MyButtonType
                Case 2 'Windows 32-bit
                    Call DrawCaption(2)
                        DrawFrame cDarkShadow, cHighLight, cShadow, cLight, False
                Case 5 'Java
                    DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, &H10), False
                    DrawRectangle 0, 0, Wi - 1, He - 1, ShiftColor(cShadow, -&H1A), True
                    DrawLine Wi - 1, 1, Wi - 1, He, cHighLight
                    DrawLine 1, He - 1, Wi - 1, He - 1, cHighLight
                    SetTextColor .hdc, cTextO
                    DrawText .hdc, elTex, Len(elTex), rc, DT_CENTER
                 Case 7, 8 'Flat buttons
                    Call DrawCaption(2)
                    DrawFrame cShadow, cHighLight, 0, 0, False, True
                Case 9 'Office XP
                    If isOver Then DrawRectangle 0, 0, Wi, He, Abs(MyColorType = 2) * ShiftColor(OXPf, -&H20) + Abs(MyColorType <> 2) * ShiftColorOXP(OXPb, &H80)
                    Call DrawCaption(2)
                    DrawRectangle 0, 0, Wi, He, OXPb, True
            End Select
            Call DrawPictures(1)
        End If
    Else
    '#~#~#~#~#~# DISABLED STATUS #~#~#~#~#~#
        Select Case MyButtonType
            Case 2 'Windows 32-bit
                Call DrawCaption(3)
                DrawFrame cHighLight, cDarkShadow, cLight, cShadow, False
            Case 5 'Java
                Call DrawCaption(4)
                DrawRectangle 0, 0, Wi, He, cShadow, True
            Case 7, 8 'Flat buttons
                Call DrawCaption(3)
                If MyButtonType = [Simple Flat] Then DrawFrame cHighLight, cShadow, 0, 0, False, True
            Case 9 'Office XP
                Call DrawCaption(4)
        End Select
        Call DrawPictures(2)
    End If
    End With
    If isOver And MyColorType = Custom Then BackC = tempCol: SetColors
End Sub

Private Sub DrawRectangle(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long, Optional OnlyBorder As Boolean = False)
    On Error Resume Next
    Dim bRECT As RECT
    Dim hBrush As Long
    bRECT.Left = X
    bRECT.Top = Y
    bRECT.Right = X + Width
    bRECT.Bottom = Y + Height
    hBrush = CreateSolidBrush(Color)
    If OnlyBorder Then
        FrameRect UserControl.hdc, bRECT, hBrush
    Else
        FillRect UserControl.hdc, bRECT, hBrush
    End If
    DeleteObject hBrush
End Sub

Private Sub DrawLine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
    On Error Resume Next
    'a fast way to draw lines
    Dim pt As POINTAPI
    Dim oldPen As Long, hPen As Long
    With UserControl
        hPen = CreatePen(PS_SOLID, 1, Color)
        oldPen = SelectObject(.hdc, hPen)
        
        MoveToEx .hdc, X1, Y1, pt
        LineTo .hdc, X2, Y2
        
        SelectObject .hdc, oldPen
        DeleteObject hPen
    End With
End Sub

Private Sub DrawFrame(ByVal ColHigh As Long, ByVal ColDark As Long, ByVal ColLight As Long, ByVal ColShadow As Long, ByVal ExtraOffset As Boolean, Optional ByVal Flat As Boolean = False)
    On Error Resume Next
    'a very fast way to draw windows-like frames
    Dim pt As POINTAPI
    Dim frHe As Long, frWi As Long, frXtra As Long
    frHe = He - 1 + ExtraOffset: frWi = Wi - 1 + ExtraOffset: frXtra = Abs(ExtraOffset)
    With UserControl
        '=============================
        If ButtonType = [Windows 32-bit] Then Call DeleteObject(SelectObject(.hdc, CreatePen(PS_SOLID, 1, ColHigh)))
        '=============================
        MoveToEx .hdc, frXtra, frHe, pt
        LineTo .hdc, frXtra, frXtra
        LineTo .hdc, frWi, frXtra
        '=============================
        If ButtonType = [Windows 32-bit] Then Call DeleteObject(SelectObject(.hdc, CreatePen(PS_SOLID, 1, ColDark)))
        '=============================
        LineTo .hdc, frWi, frHe
        LineTo .hdc, frXtra - 1, frHe
        MoveToEx .hdc, frXtra + 1, frHe - 1, pt
        If Flat Then Exit Sub
        '=============================
        If ButtonType = [Windows 32-bit] Then Call DeleteObject(SelectObject(.hdc, CreatePen(PS_SOLID, 1, ColLight)))
        '=============================
        LineTo .hdc, frXtra + 1, frXtra + 1
        LineTo .hdc, frWi - 1, frXtra + 1
        '=============================
        If ButtonType = [Windows 32-bit] Then Call DeleteObject(SelectObject(.hdc, CreatePen(PS_SOLID, 1, ColShadow)))
        '=============================
        LineTo .hdc, frWi - 1, frHe - 1
        LineTo .hdc, frXtra, frHe - 1
    End With
End Sub

Private Sub mSetPixel(ByVal X As Long, ByVal Y As Long, ByVal Color As Long)
    On Error Resume Next
    Call SetPixel(UserControl.hdc, X, Y, Color)
End Sub

Private Sub SetColors()
    On Error Resume Next
    'this function sets the colors taken as a base to build
    'all the other colors and styles.
    
    If MyColorType = Custom Then
        cFace = ConvertFromSystemColor(BackC)
        cFaceO = ConvertFromSystemColor(BackO)
        cText = ConvertFromSystemColor(ForeC)
        cTextO = ConvertFromSystemColor(ForeO)
        cShadow = ShiftColor(cFace, -&H40)
        cLight = ShiftColor(cFace, &H1F)
        cHighLight = ShiftColor(cFace, &H2F) 'it should be 3F but it looks too lighter
        cDarkShadow = ShiftColor(cFace, -&HC0)
        OXPb = ShiftColor(cFace, -&H80)
        OXPf = cFace
    Else
    'if MyColorType is 1 or has not been set then use windows colors
        cFace = GetSysColor(COLOR_BTNFACE)
        cFaceO = cFace
        cShadow = GetSysColor(COLOR_BTNSHADOW)
        cLight = GetSysColor(COLOR_BTNLIGHT)
        cDarkShadow = GetSysColor(COLOR_BTNDKSHADOW)
        cHighLight = GetSysColor(COLOR_BTNHIGHLIGHT)
        cText = GetSysColor(COLOR_BTNTEXT)
        cTextO = cText
        OXPb = GetSysColor(COLOR_HIGHLIGHT)
        OXPf = ShiftColorOXP(OXPb)
    End If
    cMask = ConvertFromSystemColor(MaskC)
    XPFace = ShiftColor(cFace, &H30, False)
End Sub

Private Sub MakeRegion()
    On Error Resume Next
    'this function creates the regions to "cut" the UserControl
    'so it will be transparent in certain areas
    
    Dim rgn1 As Long, rgn2 As Long
        
        DeleteObject rgnNorm
        rgnNorm = CreateRectRgn(0, 0, Wi, He)
        rgn2 = CreateRectRgn(0, 0, 0, 0)
        
    Select Case MyButtonType
        Case 1, 5 'Windows 16-bit, Java
            rgn1 = CreateRectRgn(0, He, 1, He - 1)
            CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(Wi, 0, Wi - 1, 1)
            CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
            DeleteObject rgn1
            If MyButtonType <> 5 Then  'the above was common code
                rgn1 = CreateRectRgn(0, 0, 1, 1)
                CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
                DeleteObject rgn1
                rgn1 = CreateRectRgn(Wi, He, Wi - 1, He - 1)
                CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
                DeleteObject rgn1
            End If
    End Select
    DeleteObject rgn2
End Sub

Private Sub SetAccessKeys()
    On Error Resume Next
    'this is a TRUE access keys parser
    'the basic rule is that if an ampersand is followed by another,
    '  a single ampersand is drawn and this is not the access key.
    '  So we continue searching for another possible access key.
    
    '   I only do a second pass because no one writes text like "Me & them & everyone"
    '   so the caption prop should be "Me && them && &everyone", this is rubbish and a
    '   search like this would only waste time
    Dim ampersandPos As Long
    
    'we first clear the AccessKeys property, and will be filled if one is found
    UserControl.AccessKeys = ""
    
    If Len(elTex) > 1 Then
        ampersandPos = InStr(1, elTex, "&", vbTextCompare)
        If (ampersandPos < Len(elTex)) And (ampersandPos > 0) Then
            If Mid$(elTex, ampersandPos + 1, 1) <> "&" Then 'if text is sonething like && then no access key should be assigned, so continue searching
                UserControl.AccessKeys = LCase$(Mid$(elTex, ampersandPos + 1, 1))
            Else 'do only a second pass to find another ampersand character
                ampersandPos = InStr(ampersandPos + 2, elTex, "&", vbTextCompare)
                If Mid$(elTex, ampersandPos + 1, 1) <> "&" Then
                    UserControl.AccessKeys = LCase$(Mid$(elTex, ampersandPos + 1, 1))
                End If
            End If
        End If
    End If
End Sub

Private Function ShiftColor(ByVal Color As Long, ByVal Value As Long, Optional isXP As Boolean = False) As Long
    On Error Resume Next
    'this function will add or remove a certain color
    'quantity and return the result
    
    Dim Red As Long, Blue As Long, Green As Long
    
    'this is just a tricky way to do it and will result in weird colors for WinXP and KDE2
    'If isSoft Then Value = Value \ 2
    
    If Not isXP Then 'for XP button i use a work-aroud that works fine
        Blue = ((Color \ &H10000) Mod &H100) + Value
    Else
        Blue = ((Color \ &H10000) Mod &H100)
        Blue = Blue + ((Blue * Value) \ &HC0)
    End If
    Green = ((Color \ &H100) Mod &H100) + Value
    Red = (Color And &HFF) + Value
    
    'a bit of optimization done here, values will overflow a
    ' byte only in one direction... eg: if we added 32 to our
    ' color, then only a > 255 overflow can occurr.
    If Value > 0 Then
        If Red > 255 Then Red = 255
        If Green > 255 Then Green = 255
        If Blue > 255 Then Blue = 255
    ElseIf Value < 0 Then
        If Red < 0 Then Red = 0
        If Green < 0 Then Green = 0
        If Blue < 0 Then Blue = 0
    End If
    
    'more optimization by replacing the RGB function by its correspondent calculation
    ShiftColor = Red + 256& * Green + 65536 * Blue
End Function

Private Function ShiftColorOXP(ByVal theColor As Long, Optional ByVal Base As Long = &HB0) As Long
    On Error Resume Next
    Dim Red As Long, Blue As Long, Green As Long
    Dim Delta As Long
    
    Blue = ((theColor \ &H10000) Mod &H100)
    Green = ((theColor \ &H100) Mod &H100)
    Red = (theColor And &HFF)
    Delta = &HFF - Base
    
    Blue = Base + Blue * Delta \ &HFF
    Green = Base + Green * Delta \ &HFF
    Red = Base + Red * Delta \ &HFF
    
    If Red > 255 Then Red = 255
    If Green > 255 Then Green = 255
    If Blue > 255 Then Blue = 255
    
    ShiftColorOXP = Red + 256& * Green + 65536 * Blue
End Function

Private Sub CalcTextRects()
    On Error Resume Next
    'this sub will calculate the rects required to draw the text
    Select Case PicPosition
        Case 0
            rc2.Left = 1 + picSZ.X: rc2.Right = Wi - 2: rc2.Top = 1: rc2.Bottom = He - 2
        Case 1
            rc2.Left = 1: rc2.Right = Wi - 2 - picSZ.X: rc2.Top = 1: rc2.Bottom = He - 2
        Case 2
            rc2.Left = 1: rc2.Right = Wi - 2: rc2.Top = 1 + picSZ.Y: rc2.Bottom = He - 2
        Case 3
            rc2.Left = 1: rc2.Right = Wi - 2: rc2.Top = 1: rc2.Bottom = He - 2 - picSZ.Y
        Case 4
            rc2.Left = 1: rc2.Right = Wi - 2: rc2.Top = 1: rc2.Bottom = He - 2
    End Select
    DrawText UserControl.hdc, elTex, Len(elTex), rc2, DT_CALCRECT Or DT_WORDBREAK
    CopyRect rc, rc2: fc.X = rc.Right - rc.Left: fc.Y = rc.Bottom - rc.Top
    Select Case PicPosition
        Case 0, 2
            OffsetRect rc, (Wi - rc.Right) \ 2, (He - rc.Bottom) \ 2
        Case 1
            OffsetRect rc, (Wi - rc.Right - picSZ.X - 4) \ 2, (He - rc.Bottom) \ 2
        Case 3
            OffsetRect rc, (Wi - rc.Right) \ 2, (He - rc.Bottom - picSZ.Y - 4) \ 2
        Case 4
            OffsetRect rc, (Wi - rc.Right) \ 2, (He - rc.Bottom) \ 2
    End Select
    CopyRect rc2, rc: OffsetRect rc2, 1, 1
    
    Call CalcPicPos 'once we have the text position we are able to calculate the pic position
End Sub

Public Sub DisableRefresh()
    On Error Resume Next
    'this is for fast button editing, once you disable the refresh,
    ' you can change every prop without triggering the drawing methods.
    ' once you are done, you call Refresh.
    isShown = False
End Sub

Public Sub Refresh()
    On Error Resume Next
    Call SetColors
    Call CalcTextRects
    isShown = True
    Call Redraw(lastStat, True)
End Sub

Private Function ConvertFromSystemColor(ByVal theColor As Long) As Long
    On Error Resume Next
    Call OleTranslateColor(theColor, 0, ConvertFromSystemColor)
End Function

Private Sub DrawCaption(ByVal State As Byte)
    On Error Resume Next
    'this code is commonly shared through all the buttons so
    ' i took it and put it toghether here for easier readability
    ' of the code, and to cut-down disk size.
    
    captOpt = State
    
    With UserControl
    Select Case State 'in this select case, we only change the text color and draw only text that needs rc2, at the end, text that uses rc will be drawn
        Case 0 'normal caption
            SetTextColor .hdc, cText
        Case 1 'hover caption
            SetTextColor .hdc, cTextO
        Case 2 'down caption
            SetTextColor .hdc, cTextO
            DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTER
        Case 3 'disabled embossed caption
            SetTextColor .hdc, cHighLight
            DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTER
            SetTextColor .hdc, cShadow
        Case 4 'disabled grey caption
            SetTextColor .hdc, cShadow
        Case 5 'WinXP disabled caption
            SetTextColor .hdc, ShiftColor(XPFace, -&H68, True)
    End Select
    'we now draw the text that is common in all the captions
    If State <> 2 Then DrawText .hdc, elTex, Len(elTex), rc, DT_CENTER
    End With
End Sub

Private Sub DrawPictures(ByVal State As Byte)
    On Error Resume Next
    If picNormal Is Nothing Then Exit Sub 'check if there is a main picture, if not then exit
    With UserControl
    Select Case State
        Case 0 'normal & hover
            If Not isOver Then
                TransBlt .hdc, picPT.X, picPT.Y, picSZ.X, picSZ.Y, picNormal, cMask, , , , (MyButtonType = [Office XP])
            Else
                If MyButtonType = [Office XP] Then
                    TransBlt .hdc, picPT.X + 1, picPT.Y + 1, picSZ.X, picSZ.Y, picNormal, cMask, cShadow
                    TransBlt .hdc, picPT.X - 1, picPT.Y - 1, picSZ.X, picSZ.Y, picNormal, cMask
                Else
                    If Not picHover Is Nothing Then
                        TransBlt .hdc, picPT.X, picPT.Y, picSZ.X, picSZ.Y, picHover, cMask
                    Else
                        TransBlt .hdc, picPT.X, picPT.Y, picSZ.X, picSZ.Y, picNormal, cMask
                    End If
                End If
            End If
        Case 1 'down
            If picHover Is Nothing Or MyButtonType = [Office XP] Then
                Select Case MyButtonType
                Case 5, 9
                    TransBlt .hdc, picPT.X, picPT.Y, picSZ.X, picSZ.Y, picNormal, cMask
                Case Else
                    TransBlt .hdc, picPT.X + 1, picPT.Y + 1, picSZ.X, picSZ.Y, picNormal, cMask
                End Select
            Else
                TransBlt .hdc, picPT.X + Abs(MyButtonType <> [Java metal]), picPT.Y + Abs(MyButtonType <> [Java metal]), picSZ.X, picSZ.Y, picHover, cMask
            End If
        Case 2 'disabled
            Select Case MyButtonType
            Case 5, 6, 9    'draw flat grey pictures
                TransBlt .hdc, picPT.X, picPT.Y, picSZ.X, picSZ.Y, picNormal, cMask, Abs(MyButtonType = [Office XP]) * ShiftColor(cShadow, &HD) + Abs(MyButtonType <> [Office XP]) * cShadow, True
            Case 3          'for WinXP draw a greyscaled image
                TransBlt .hdc, picPT.X + 1, picPT.Y + 1, picSZ.X, picSZ.Y, picNormal, cMask, , , True
            Case Else       'draw classic embossed pictures
                TransBlt .hdc, picPT.X + 1, picPT.Y + 1, picSZ.X, picSZ.Y, picNormal, cMask, cHighLight, True
                TransBlt .hdc, picPT.X, picPT.Y, picSZ.X, picSZ.Y, picNormal, cMask, cShadow, True
            End Select
    End Select
    End With
    If PicPosition = cbBackground Then Call DrawCaption(captOpt)
End Sub

Private Sub CalcPicSize()
    On Error Resume Next
    If Not picNormal Is Nothing Then
        picSZ.X = UserControl.ScaleX(picNormal.Width, 8, UserControl.ScaleMode)
        picSZ.Y = UserControl.ScaleY(picNormal.Height, 8, UserControl.ScaleMode)
    Else
        picSZ.X = 0: picSZ.Y = 0
    End If
End Sub

Private Sub CalcPicPos()
    On Error Resume Next
    'exit if there's no picture
    If picNormal Is Nothing And picHover Is Nothing Then Exit Sub
    
    If (Trim$(elTex) <> "") And (PicPosition <> 4) Then 'if there is no caption, or we have the picture as background, then we put the picture at the center of the button
        Select Case PicPosition
            Case 0 'left
                picPT.X = rc.Left - picSZ.X - 4
                picPT.Y = (He - picSZ.Y) \ 2
            Case 1 'right
                picPT.X = rc.Right + 4
                picPT.Y = (He - picSZ.Y) \ 2
            Case 2 'top
                picPT.X = (Wi - picSZ.X) \ 2
                picPT.Y = rc.Top - picSZ.Y - 2
            Case 3 'bottom
                picPT.X = (Wi - picSZ.X) \ 2
                picPT.Y = rc.Bottom + 2
        End Select
    Else 'center the picture
        picPT.X = (Wi - picSZ.X) \ 2
        picPT.Y = (He - picSZ.Y) \ 2
    End If
End Sub

Private Sub TransBlt(ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcPic As StdPicture, Optional ByVal TransColor As Long = -1, Optional ByVal BrushColor As Long = -1, Optional ByVal MonoMask As Boolean = False, Optional ByVal isGreyscale As Boolean = False, Optional ByVal XPBlend As Boolean = False)
    On Error Resume Next
    If DstW = 0 Or DstH = 0 Then Exit Sub
        
    Dim B As Long, H As Long, F As Long, I As Long, newW As Long
    Dim TmpDC As Long, TmpBmp As Long, TmpObj As Long
    Dim Sr2DC As Long, Sr2Bmp As Long, Sr2Obj As Long
    Dim Data1() As RGBTRIPLE, Data2() As RGBTRIPLE
    Dim Info As BITMAPINFO, BrushRGB As RGBTRIPLE, gCol As Long
    
    Dim SrcDC As Long, tObj As Long, ttt As Long

    SrcDC = CreateCompatibleDC(hdc)
    
    If DstW < 0 Then DstW = UserControl.ScaleX(SrcPic.Width, 8, UserControl.ScaleMode)
    If DstH < 0 Then DstH = UserControl.ScaleY(SrcPic.Height, 8, UserControl.ScaleMode)
    
    If SrcPic.Type = 1 Then 'check if it's an icon or a bitmap
        tObj = SelectObject(SrcDC, SrcPic)
    Else
        Dim br As RECT, hBrush As Long: br.Right = DstW: br.Bottom = DstH
        ttt = CreateCompatibleBitmap(DstDC, DstW, DstH): tObj = SelectObject(SrcDC, ttt)
        hBrush = CreateSolidBrush(MaskColor): FillRect SrcDC, br, hBrush
        DeleteObject hBrush
        DrawIconEx SrcDC, 0, 0, SrcPic.handle, 0, 0, 0, 0, &H1 Or &H2
    End If
 
    TmpDC = CreateCompatibleDC(SrcDC)
    Sr2DC = CreateCompatibleDC(SrcDC)
    TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    TmpObj = SelectObject(TmpDC, TmpBmp)
    Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
    ReDim Data1(DstW * DstH * 3 - 1)
    ReDim Data2(UBound(Data1))
    With Info.bmiHeader
        .biSize = Len(Info.bmiHeader)
        .biWidth = DstW
        .biHeight = DstH
        .biPlanes = 1
        .biBitCount = 24
    End With
    
    BitBlt TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
    BitBlt Sr2DC, 0, 0, DstW, DstH, SrcDC, 0, 0, vbSrcCopy
    GetDIBits TmpDC, TmpBmp, 0, DstH, Data1(0), Info, 0
    GetDIBits Sr2DC, Sr2Bmp, 0, DstH, Data2(0), Info, 0
    
    If BrushColor > 0 Then
        BrushRGB.rgbBlue = (BrushColor \ &H10000) Mod &H100
        BrushRGB.rgbGreen = (BrushColor \ &H100) Mod &H100
        BrushRGB.rgbRed = BrushColor And &HFF
    End If
    
    'If Not useMask Then TransColor = -1
    
    newW = DstW - 1
    
    For H = 0 To DstH - 1
        F = H * DstW
        For B = 0 To newW
            I = F + B
            If (CLng(Data2(I).rgbRed) + 256& * Data2(I).rgbGreen + 65536 * Data2(I).rgbBlue) <> TransColor Then
            With Data1(I)
                If BrushColor > -1 Then
                    If MonoMask Then
                        If (CLng(Data2(I).rgbRed) + Data2(I).rgbGreen + Data2(I).rgbBlue) <= 384 Then Data1(I) = BrushRGB
                    Else
                        Data1(I) = BrushRGB
                    End If
                Else
                    If isGreyscale Then
                        gCol = CLng(Data2(I).rgbRed * 0.3) + Data2(I).rgbGreen * 0.59 + Data2(I).rgbBlue * 0.11
                        .rgbRed = gCol: .rgbGreen = gCol: .rgbBlue = gCol
                    Else
                        If XPBlend Then
                            .rgbRed = (CLng(.rgbRed) + Data2(I).rgbRed * 2) \ 3
                            .rgbGreen = (CLng(.rgbGreen) + Data2(I).rgbGreen * 2) \ 3
                            .rgbBlue = (CLng(.rgbBlue) + Data2(I).rgbBlue * 2) \ 3
                        Else
                            Data1(I) = Data2(I)
                        End If
                    End If
                End If
            End With
            End If
        Next
    Next

    SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data1(0), Info, 0

    Erase Data1, Data2
    DeleteObject SelectObject(TmpDC, TmpObj)
    DeleteObject SelectObject(Sr2DC, Sr2Obj)
    If SrcPic.Type = 3 Then DeleteObject SelectObject(SrcDC, tObj)
    DeleteDC TmpDC: DeleteDC Sr2DC
    DeleteObject tObj: DeleteObject ttt: DeleteDC SrcDC
End Sub

Private Function isMouseOver() As Boolean
    On Error Resume Next
    Dim pt As POINTAPI
    GetCursorPos pt
    isMouseOver = (WindowFromPoint(pt.X, pt.Y) = hWnd)
End Function
