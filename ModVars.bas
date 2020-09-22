Attribute VB_Name = "ModVars"
Option Explicit

'DEFINITION
'VAR: VARIABLES
'PARAM: PARAMETERS USED IN CURRENT LINE

Public Enum eType
    eVariant = 0
    eString = 1
    eInteger = 2
    eLong = 3
    eDouble = 4
    eDate = 5
    eBool = 6
End Enum

Public Type JustPointers
    jName As String
    jLineNumber As Long
End Type

Public Type VarPointers
    vName As String
    vType As eType
    vValue As Variant
End Type

Public Type ArrayPointers
    aName As String
    aValue As String
    aSeparator As String
End Type

Public myParams() As String
Public myPointers() As JustPointers '1000 ought to be enough for everybody - more than enough for 2008 man.
Public myVars() As VarPointers '1000 ought to be enough for everybody - more than enough for 2008 man.
Public myArrays() As ArrayPointers


'Public subStopper As Boolean 'this is true if the sub has to be stopped

'returns line number when name is given
Function Pointer(ByVal pName As String, Optional ByVal Default = 10000) As Long
    On Error Resume Next
    Dim K As String
    Dim I As Long
    Dim HoaxMark As Boolean
    Pointer = -1 'assign dead value
    
'    SDebug "Pointer is run", 2
    K = LCase$(pName)
    If Len(K) = 0 Then HoaxMark = True 'if you didnt even tell me what to find
    
    If HoaxMark = False Then
        For I = LBound(myPointers) To UBound(myPointers) Step 1
            If LCase$(myPointers(I).jName) = K Then
                'If myPointers(I).jLineNumber <> 0 Then
                    Pointer = myPointers(I).jLineNumber
                    SDebug "Pointer seek of " & K & " returns line " & Pointer, 0
                'End If
                Exit For 'stop searching
            End If
        Next
'    Else
'        SDebug "HoaxMark is true - pointer not assigned", 2
    End If
    If Pointer = -1 Then Pointer = Default
End Function

Function AssignPointer(ByVal pName As String, ByVal pLineNumber As Long)
    Dim K As String
    Dim I As Long
    K = LCase$(pName)
    
    For I = LBound(myPointers) To UBound(myPointers) Step 1
        If myPointers(I).jName = K Then
                myPointers(I).jLineNumber = pLineNumber
                Exit For 'stop looping
        End If
    Next
    If I > UBound(myPointers) Then 'if the name is not found
        For I = LBound(myPointers) To UBound(myPointers) Step 1
            If myPointers(I).jName = "" Then
                    myPointers(I).jName = K
                    myPointers(I).jLineNumber = pLineNumber
                    Exit For 'stop looping
            End If
        Next
    End If
End Function

Function Var(ByVal valName As String, Optional ByVal Default = "")
    Dim I As Long
    Dim HoaxMark As Boolean
    
    Var = vbNullString
    
    If Len(valName) = 0 Then HoaxMark = True 'if you didnt even tell me what to find
    
    If HoaxMark = False Then
        For I = LBound(myVars) To UBound(myVars) Step 1
            If LCase$(myVars(I).vName) = LCase$(valName) Then
                If myVars(I).vValue <> vbNullString Then
                    Var = myVars(I).vValue
                    Select Case myVars(I).vType
                        Case 1 '    eString = 1
                            Var = CStr(Var)
                        Case 2 '    eInteger = 2
                            Var = CInt(Var)
                        Case 3 '    eLong = 3
                            Var = CLng(Var)
                        Case 4 '    eDouble = 4
                            Var = CDbl(Var)
                        Case 5 '    eDate = 5
                            Var = CDate(Var)
                        Case 6
                            Var = CBool(Var)
                        Case Else
                            'Var = CVar(Var) 'what a waste of time, but safer?
                            DoEvents 'doesnt this make variant the fastest variable...
                    End Select
                End If
                Exit For 'stop searching
            End If
        Next
    End If
    If Var = vbNullString Then Var = Default
End Function

Function AssignVar(valName As String, valValue, Optional VarType As eType)
    Dim I As Long
    For I = LBound(myVars) To UBound(myVars) Step 1
        If myVars(I).vName = valName Then
                myVars(I).vValue = valValue
                Exit For 'stop looping
        End If
    Next
    If I > UBound(myVars) Then 'if the name is not found
        For I = LBound(myVars) To UBound(myVars) Step 1
            If myVars(I).vName = "" Then
                    myVars(I).vName = valName
                    myVars(I).vType = VarType
                    myVars(I).vValue = valValue
                    Exit For 'stop looping
            End If
        Next
    End If
    SDebug "var " & valName & " = " & valValue, 0
End Function

Function Params(Index As Integer, Optional Default = "") As Variant
    Dim HoaxMark As Boolean
    If Index > UBound(myParams) Or Index < LBound(myParams) Then 'if its existance is not even possible then
        HoaxMark = True
        GoTo SkipContentValidation
    End If
    On Error GoTo SkipContentValidation
    If myParams(Index) = "" Then HoaxMark = True 'if the variable has no value
SkipContentValidation:
    If HoaxMark Then
        Params = Default
    Else
        Params = myParams(Index)
    End If
    
    Params = SearchAndReplaceParamsWithVals(CStr(Params))
    Params = UserFriendly(Params)
End Function

Function SearchAndReplaceParamsWithVals(ByVal What As String) As String
    On Error Resume Next
    'replacing variables
    Dim varSearchStart As Long, varSearchEnd As Long
    Dim varSearchFound As String
    Dim BFR As String 'buffer
    BFR = What
    Do
        varSearchStart = InStr(1, BFR, "[", vbTextCompare)
        If varSearchStart > 0 Then 'if signs suggest that there's a var callback
            varSearchEnd = InStr(varSearchStart, BFR, "]", vbTextCompare) 'then find the ending bit
            If varSearchEnd > 0 Then
                varSearchEnd = varSearchEnd - varSearchStart
            Else
                Exit Do 'there seems to be nothing left to replace
            End If
            varSearchFound = Mid$(BFR, varSearchStart + 1, varSearchEnd - 1) '+4 -5 to remove var( and )

            BFR = Replace(BFR, "[" & varSearchFound & "]", Var(varSearchFound), , , vbTextCompare)
            If Len(varSearchFound) > 0 Then
                varSearchStart = 0
                varSearchEnd = 0
                varSearchFound = vbNullString
            End If
        End If
    Loop Until InStr(1, BFR, "[", vbTextCompare) = 0 'until theres nothing left to replace
    'replacing variables end
    SearchAndReplaceParamsWithVals = BFR

'    On Error Resume Next
'    'replacing variables
'    Dim varSearchStart As Long, varSearchEnd As Long
'    Dim varSearchFound As String
'    Dim BFR As String 'buffer
'    BFR = What
'    Do
'        varSearchStart = InStr(1, BFR, "[", vbTextCompare)
'        If varSearchStart > 0 Then 'if signs suggest that there's a var callback
'            varSearchEnd = InStr(varSearchStart, BFR, "]") 'then find the ending bit
'            If varSearchEnd > 0 Then varSearchEnd = varSearchEnd - varSearchStart
'            varSearchFound = Mid$(BFR, varSearchStart + 1, varSearchEnd - 1) '+4 -5 to remove var( and )
'
'            If Len(varSearchFound) > 0 Then
'                BFR = Replace(BFR, "[" & varSearchFound & "]", Var(varSearchFound), , , vbTextCompare)
'                varSearchStart = 0
'                varSearchEnd = 0
'                varSearchFound = vbNullString
'            Else
'                SDebug "Attempt to request value of zero length variable was blocked: " & varSearchFound, 0
'            End If
'        End If
'    Loop Until InStr(1, BFR, "[", vbTextCompare) = 0 'until theres nothing left to replace
'    'replacing variables end
'    SearchAndReplaceParamsWithVals = BFR
End Function

Function UserFriendly(ByVal BFR As String) As String
    If InStr(1, BFR, "\n", vbTextCompare) > 0 Then BFR = Replace(BFR, "\n", vbCrLf, , , vbTextCompare) ' RETURN
    If InStr(1, BFR, "\cma", vbTextCompare) > 0 Then BFR = Replace(BFR, "\cma", ",", , , vbTextCompare) ' ,
    If InStr(1, BFR, "\smc", vbTextCompare) > 0 Then BFR = Replace(BFR, "\smc", ";", , , vbTextCompare) ' ;
    If InStr(1, BFR, "\obr", vbTextCompare) > 0 Then BFR = Replace(BFR, "\obr", "(", , , vbTextCompare) ' (
    If InStr(1, BFR, "\cbr", vbTextCompare) > 0 Then BFR = Replace(BFR, "\cbr", ")", , , vbTextCompare) ' )
    If InStr(1, BFR, "\osb", vbTextCompare) > 0 Then BFR = Replace(BFR, "\osb", "[", , , vbTextCompare) ' [
    If InStr(1, BFR, "\csb", vbTextCompare) > 0 Then BFR = Replace(BFR, "\csb", "]", , , vbTextCompare) ' ]
    If InStr(1, BFR, "\eql", vbTextCompare) > 0 Then BFR = Replace(BFR, "\eql", "=", , , vbTextCompare) ' =
    If InStr(1, BFR, "\\", vbTextCompare) > 0 Then BFR = Replace(BFR, "\\", "\", , , vbTextCompare) ' \
    If InStr(1, BFR, "\t", vbTextCompare) > 0 Then BFR = Replace(BFR, "\t", "    ", , , vbTextCompare) ' TAB
    UserFriendly = BFR
End Function

Function CodeFriendly(ByVal BFR As String) As String
    BFR = Replace(BFR, vbCrLf, "")
    If InStr(1, BFR, "\,") > 0 Then BFR = Replace(BFR, "\,", "\cma") ' ,
    If InStr(1, BFR, "\;") > 0 Then BFR = Replace(BFR, "\;", "\smc") ' ;
    If InStr(1, BFR, "\(") > 0 Then BFR = Replace(BFR, "\(", "\obr") ' (
    If InStr(1, BFR, "\)") > 0 Then BFR = Replace(BFR, "\)", "\cbr") ' )
    If InStr(1, BFR, "\[") > 0 Then BFR = Replace(BFR, "\[", "\osb") ' [
    If InStr(1, BFR, "\]") > 0 Then BFR = Replace(BFR, "\]", "\csb") ' ]
    If InStr(1, BFR, "\=") > 0 Then BFR = Replace(BFR, "\=", "\eql") ' =
    If InStr(1, BFR, vbTab, vbTextCompare) > 0 Then BFR = Replace(BFR, vbTab, "", , , vbTextCompare) ' TAB
    CodeFriendly = BFR
End Function

