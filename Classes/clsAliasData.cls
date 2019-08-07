VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAliasData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Hostmask As String

Private collAliases As Collection

Private Sub Class_Initialize()
    Set collAliases = New Collection
End Sub

Public Sub AddAlias(ByVal Alias As String)
    Dim tempString As Variant
    For Each tempString In collAliases
        If Alias = tempString Then
            Exit Sub
        End If
    Next

    collAliases.Add Alias
End Sub

Public Function GetAliases() As String
    Dim tempString As Variant
    For Each tempString In collAliases
        GetAliases = GetAliases & tempString & "."
    Next
    GetAliases = Left$(GetAliases, Len(GetAliases) - 1)
End Function

