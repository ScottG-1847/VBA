Attribute VB_Name = "GUIDs"
Option Explicit

'Easier to code
Public Function GetGUID() As String
    GetGUID = Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36) 'Just the GUID
    'GetGUID = Left$(CreateObject("Scriptlet.TypeLib").Guid, 38)   'GUID w/ {}'s
End Function

'==========================

'About 10 times faster, but they're both really fast in human terms
Private Type UUID
    Data1    As Long
    Data2    As Integer
    Data3    As Integer
    Data4(7) As Byte
End Type

Private Declare Function UuidCreate Lib "rpcrt4.dll" ( _
                                    ByRef lpUuid As Any _
                                ) As Long

Private Declare Function StringFromGUID2 Lib "ole32.dll" ( _
                                    ByRef rguid As Any, _
                                    ByVal lpsz As Long, _
                                    ByVal cchMax As Long _
                                ) As Long
                                
Private Declare Function SysReAllocStringLen Lib "oleaut32.dll" ( _
                                    ByVal pBSTR As Long, _
                                    Optional ByVal pszStrPtr As Long, _
                                    Optional ByVal Length As Long _
                                ) As Long

Public Function CreateGUID() As String
    Const RPC_S_OK = 0&
    
    Dim udtUUID As UUID

    If UuidCreate(udtUUID) = RPC_S_OK Then
        SysReAllocStringLen VarPtr(CreateGUID), , 38&: _
        StringFromGUID2 udtUUID, StrPtr(CreateGUID), 39&
    End If
End Function


