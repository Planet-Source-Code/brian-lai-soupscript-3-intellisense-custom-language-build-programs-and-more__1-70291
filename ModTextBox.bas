Attribute VB_Name = "ModTextBox"
'[ModTextBox]
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long

    Private Const EM_GETLINECOUNT = &HBA
    Private Const EM_LINEINDEX = &HBB
    Private Const EM_LINELENGTH = &HC1
    Private Const EM_GETLINE = &HC4
    Private Const EM_LINEFROMCHAR = &HC9
    Private Const EM_POSFROMCHAR = &HD6
'Thanks a PSC guy for teaching me this

Public Function GetTextboxLine(txtTextbox As Control, Optional ByVal lngLineNumber As Long = 1) As String
'    Dim intLineLength As Integer: Dim strLine As String
'    Dim lngIndexChar As Long, lngRetValue As Long, lngLineCount As Long
'    lngLineCount = SendMessage(txtTextbox.hWnd, EM_GETLINECOUNT, 0&, ByVal 0&)    'get total number of lines
'    If lngLineNumber > lngLineCount Or lngLineNumber < 1 Then    'If the line requested is beyond the end or before the beginning of the text box
'        GetTextboxLine = "": Exit Function       'return an empty string
'    End If
'    lngLineNumber = lngLineNumber - 1    'subtract 1 from requested line number because the API uses 0 based line numbers
'    lngIndexChar = SendMessage(txtTextbox.hWnd, EM_LINEINDEX, ByVal lngLineNumber, ByVal 0&)    'get the text position of the first char in requested line
'    intLineLength = SendMessage(txtTextbox.hWnd, EM_LINELENGTH, ByVal lngIndexChar, ByVal 0&)    'get the length of that line
'    strLine = Space$(IIf(intLineLength >= 2, intLineLength, 2))    'fill a buffer string with spaces- minimum length of 2 required for the next line
'    CopyMemory ByVal strLine, intLineLength, Len(intLineLength)    'copy the binary value of linelength into the beginning of the buffer string this is necessary becuase the last argument for this call needs to contain the length and is also used to return the string.
'    lngRetValue = SendMessage(txtTextbox.hWnd, EM_GETLINE, ByVal lngLineNumber, ByVal strLine)    'put the text of the line into the string buffer - strLine contains the length
'    If intLineLength = 1 Then strLine = Left$(strLine, 1)    'trim ending null char if the line is only 1 char long
'    GetTextboxLine = strLine    'return the line
    Dim lngRet As Long
    Dim lngLen As Long
    Dim lngFirstCharPos As Long
    Dim lngHwnd As Long
    Dim bytBuffer() As Byte
    Dim strAns As String
    
    If lngLineNumber < 0 Then Exit Function
    
    If txtTextbox.MultiLine = False Then
        GetTextboxLine = txtTextbox.Text
    Else
        lngHwnd = txtTextbox.hWnd
        'first character position of the line
        lngFirstCharPos = SendMessage(lngHwnd, EM_LINEINDEX, lngLineNumber - 1, 0&)
        'length of line
        lngLen = SendMessage(lngHwnd, EM_LINELENGTH, lngFirstCharPos, 0&)
        
        ReDim bytBuffer(lngLen) As Byte
        bytBuffer(0) = lngLen
        
        'text of line saved to bytBuffer
        lngRet = SendMessage(lngHwnd, EM_GETLINE, lngLineNumber - 1, bytBuffer(0))
        If lngRet > 0 Then
            strAns = Left$(StrConv(bytBuffer, vbUnicode), lngLen)
        End If
        GetTextboxLine = strAns
    End If
End Function

Public Function GetLineCount(MyTextBox As TextBox) As Long
    GetLineCount = SendMessage(MyTextBox.hWnd, EM_GETLINECOUNT, 0, ByVal 0&)    'returns number of lines in a multiline textbox
End Function

Public Function GetLinePos(MyTextBox As TextBox) As Long
    GetLinePos = SendMessage(MyTextBox.hWnd, EM_LINEFROMCHAR, -1, ByVal 0&) + 1
End Function

Public Function CaretX(MyTextBox As TextBox) As Long
    Dim P As POINTAPI
    GetCaretPos P
    CaretX = P.X * Screen.TwipsPerPixelX
End Function

Public Function CaretY(MyTextBox As TextBox) As Long
    Dim P As POINTAPI
    GetCaretPos P
    CaretY = P.Y * Screen.TwipsPerPixelY
End Function
