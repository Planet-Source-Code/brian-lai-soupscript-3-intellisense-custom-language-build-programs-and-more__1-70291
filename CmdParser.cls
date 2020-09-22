VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCmdParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Option Explicit
Dim myCode As String
Dim myLineNumber As Long
Dim myACC
Dim myReturnVal As Variant
Dim ExplicitGoToLine As Long 'this is >= 0 when the next line is different from normal... like caused by a jump.

Property Get Code() As String
    Code = myCode
End Property

Property Let Code(What As String)
    myCode = What
End Property

Property Let AssignedLineNumber(What As Long)
    myLineNumber = What
End Property

Property Get ReturnValue() 'As Variant
    ReturnValue = myReturnVal
End Property

Property Get NextLineVal() As Long
    On Error Resume Next
    NextLineVal = ExplicitGoToLine
End Property

Property Let NextLineVal(What As Long)
    On Error Resume Next
    ExplicitGoToLine = What
End Property

Function RunCode()
    On Error Resume Next
    Dim CmdStart As Long, CmdEnd As Long, CmdEnd2 As Long, cmd As String
    Dim ParamsStart As Long, ParamsEnd As Long, ParamsBuffer As String
    Dim AssStart As Long, AssEnd As Long, Assignment As String
    Dim IsAssumedImplicit As Boolean
    Dim IsAllCmd As Long
    Dim EN As Long, ED As String 'errors
    
    AssEnd = InStr(1, myCode, "=")
    Assignment = Mid$(myCode, 1, IIf(AssEnd = 0, 0, AssEnd - 1))
    Assignment = Trim$(Assignment) 'just to make sure
    
    CmdStart = IIf(Len(Assignment) > 0, AssEnd + 1, 1) 'probably cos the command starts at the beginning?
    CmdEnd = InStr(CmdStart, myCode, "(")
    CmdEnd2 = InStr(CmdStart, myCode, " ")
    If CmdEnd = 0 Then IsAssumedImplicit = True
    If CmdEnd2 < CmdEnd And CmdEnd2 <> 0 Then IsAssumedImplicit = True
    If IsAssumedImplicit Then CmdEnd = CmdEnd2
    
    IsAllCmd = CmdEnd - CmdStart
    If IsAllCmd < 0 Then 'if the whole line is a command
        cmd = Mid$(myCode, CmdStart)
    Else 'if there are commands
        cmd = Mid$(myCode, CmdStart, CmdEnd - CmdStart)
    End If
    
    If Len(cmd) > 0 Then 'if something needs to be executed
        If IsAllCmd < 0 Then 'if the whole line is a command
            ParamsBuffer = ""
            ReDim myParams(0) As String 'erase param contents. you wont be needing any
        Else 'if there are commands
            ParamsStart = CmdEnd 'probably cos the params follow right after?
            ParamsEnd = Len(myCode) 'probably cos the line ends?
            ParamsBuffer = Mid$(myCode, ParamsStart + 1, ParamsEnd - ParamsStart - IIf(IsAssumedImplicit, 0, 1))
            If Len(ParamsBuffer) > 0 Then myParams = Split(ParamsBuffer, ",")
        End If

        Dim ThisLine As New ClsFunctions
        ThisLine.AssignedLineNumber = myLineNumber
        ThisLine.AllmyParams = ParamsBuffer 'let the code know all the params just in case it needs the complete set, like "RunCode"
        CallByName ThisLine, cmd & "0", VbMethod
        
        EN = Err.Number
        ED = Err.Description
        
        myReturnVal = ThisLine.ReturnValue
        ExplicitGoToLine = ThisLine.NextLineVal 'just a way of passing it back to the superclass...
        Set ThisLine = Nothing
        If Len(Assignment) > 0 Then
            'SDebug "Error code:" & EN & "; " & ED, 2
            If EN = 438 Then 'command not found
                SDebug "The whole command " & Mid$(myCode, AssEnd + 1) & " has been allocated into variable", 0
                AssignVar Assignment, Mid$(myCode, AssEnd + 1) 'assign anything behind the equal signs
            Else
                AssignVar Assignment, myReturnVal
                If EN <> 0 Then
                    SDebug "ERROR " & EN & ": " & ED, 2
                End If
            End If
        End If
    End If
End Function

Private Sub Class_Initialize()
     'this thing is a requirement so that the use of this var below does not trigger double calls
    ExplicitGoToLine = -1 'as to produce "something hasn't really happened" long
End Sub
