VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFunctionParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim myCode As String
Dim myLineNumber As Long
Dim myReturnVal As Variant
Dim myACC As String
Dim ExplicitGoToLine As Long
Dim FunctionLineStart As Long, FunctionLineEnd As Long

Property Get Code() As String
    Code = myCode
End Property

Property Let Code(What As String)
    myCode = What
End Property

Property Get ACC()
    ACC = myACC
End Property

Property Let ACC(What)
    myACC = What
End Property

Property Let FunctionStart(What As Long)
    FunctionLineStart = What
End Property

Property Get FunctionStart() As Long
    FunctionStart = FunctionLineStart
End Property

Property Let FunctionEnd(What As Long)
    FunctionLineEnd = What
End Property

Property Get FunctionEnd() As Long
    FunctionEnd = FunctionLineEnd
End Property

Property Get ReturnValue() 'As Variant
    ReturnValue = myReturnVal
End Property

Function RunCode()
    Dim cmd As New ClsCmdParser 'create new hive
    Dim Lines() As String, CurrentLine As String
    Dim LineReturn
    Dim Looper As Long
    Lines = Split(myCode, ";")
    
    'private looper for handling "mark" in advance
    
    If FunctionLineStart = 0 Then FunctionLineStart = LBound(Lines) 'define boundary
    If FunctionLineEnd = 0 Then FunctionLineEnd = UBound(Lines) 'define boundary
    
    For Looper = FunctionLineStart To FunctionLineEnd Step 1
        DoEvents
        CurrentLine = Lines(Looper)
        'CurrentLine = Replace(CurrentLine, vbCrLf, vbNullString) 'make an enter "nothing"
        If Len(CurrentLine) > 0 Then
            If LCase$(Left$(CurrentLine, 4)) = "mark" _
            Or LCase$(Left$(CurrentLine, 3)) = "end" _
            Or LCase$(Left$(CurrentLine, 3)) = "sub" Then
            'Or LCase$(Left$(CurrentLine, 5)) = "begin" Then 'prioritized keywords
                cmd.AssignedLineNumber = Looper
                cmd.Code = CurrentLine
                cmd.RunCode
            End If
            Set cmd = Nothing
        End If
    Next
    
    For Looper = FunctionLineStart To FunctionLineEnd Step 1
        DoEvents
        CurrentLine = Lines(Looper)
        'CurrentLine = Replace(CurrentLine, vbCrLf, vbNullString) 'make an enter "nothing"
        If Len(CurrentLine) > 0 Then
            'Dim cmd As New ClsCmdParser 'create new hive
            cmd.AssignedLineNumber = Looper
'            SDebug CurrentLine, 0
            cmd.Code = CurrentLine
            cmd.RunCode
            LineReturn = cmd.ReturnValue
            SDebug "-------------------------" & vbCrLf & _
                        CurrentLine & " ->""" & First(CStr(LineReturn), 10) & """" & vbCrLf & _
                        "Position: " & Looper & " (" & _
                        FunctionLineStart & "-" & _
                        FunctionLineEnd & ")", 0
            ExplicitGoToLine = cmd.NextLineVal
            If ExplicitGoToLine >= 0 Then
                Looper = ExplicitGoToLine - 1 ' - 1 i think because looper +1 after this sub
                ExplicitGoToLine = -1 'restore yourself - nobody cleans up your ass otherwise
            End If
        End If
        Set cmd = Nothing
    Next
    myReturnVal = LineReturn
End Function

Private Sub Class_Initialize()
    ExplicitGoToLine = -1
End Sub
