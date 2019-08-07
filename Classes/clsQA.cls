VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private collQA As Collection

Public Sub SaveData()
    Dim n As Integer
    n = FreeFile
    Open "qa.txt" For Output As #n
        Dim strTemp As Variant
        For Each strTemp In collQA
            Print #n, strTemp
        Next
    Close #n
End Sub

Public Sub LoadData()
    Dim n As Integer
    n = FreeFile
    Open "qa.txt" For Input As #n
        Dim strQ As String, strA As String, strTemp As String
        Do While Not EOF(n)
            Line Input #n, strTemp
            strQ = Split(strTemp, "|")(0)
            strA = Split(strTemp, "|")(1)
            AddQAPair strQ, strA
        Loop
    Close #n
End Sub

Public Sub AddQAPair(ByVal Question As String, ByVal Answer As String)
    collQA.Add Question & "|" & Answer, Question
End Sub

Public Function CheckQA(ByVal inString As String) As String
    Dim strTemp As Variant, strQ As String, strA As String
    For Each strTemp In collQA
        strQ = Split(strTemp, "|")(0)
        strA = Split(strTemp, "|")(1)
        
        'Debug.Print LCase$(inString) & " / " & LCase$(strQ)
        If LCase$(inString) Like LCase$(strQ) Then
            CheckQA = strA
        End If
    Next
End Function

Public Function MatchQA(ByVal inString As String) As String
    Dim strTemp As Variant, strQ As String, strA As String
    For Each strTemp In collQA
        strQ = Split(strTemp, "|")(0)
        strA = Split(strTemp, "|")(1)
        
        'Debug.Print LCase$(inString) & " / " & LCase$(strQ)
        If LCase$(inString) Like LCase$(strQ) Then
            MatchQA = strQ
        End If
    Next
End Function

Public Function RemoveQA(ByVal inString As String)
    collQA.Remove (inString)
End Function

Private Sub Class_Initialize()
    Set collQA = New Collection
End Sub
