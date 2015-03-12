Attribute VB_Name = "HasValue_NoValue"
Option Explicit

Public Function HasValue(v As Variant) As Boolean
    If NoValue(v) Then
        HasValue = False
    Else
        HasValue = True
    End If
End Function

Public Function NoValue(v As Variant) As Boolean
    If IsNull(v) Then
        NoValue = True
    ElseIf Len(CStr(v)) = 0 Then
        NoValue = True
    ElseIf (Trim(CStr(v))) = "" Then
        NoValue = True
    Else
        NoValue = False
    End If
End Function
