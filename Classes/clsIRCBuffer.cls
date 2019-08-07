VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIRCBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strData As String

Public Sub AddData(ByVal Data As String)
    strData = strData & Data
End Sub

Public Function IsFullPacket() As Boolean
    If InStr(strData, vbCrLf) > 0 Then
        IsFullPacket = True
    End If
End Function

Public Function GetPacket() As String
    Dim tempPos As Integer
    tempPos = InStr(strData, vbCrLf)
    GetPacket = Mid$(strData, 1, tempPos - 1)
    strData = Mid$(strData, tempPos + 2)
End Function
