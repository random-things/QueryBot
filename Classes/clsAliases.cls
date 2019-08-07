VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAliases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private collAliases As Collection

Public Sub SaveData()
    'On Error Resume Next
    Dim n As Integer
    n = FreeFile
    Open "aliases.txt" For Output As #n
        Dim tempAlias As clsAliasData, tempString As String
        For Each tempAlias In collAliases
            tempString = tempAlias.Hostmask & "/" & tempAlias.GetAliases
            Print #n, tempString
        Next
    Close #n
End Sub

Public Sub LoadData()
    Dim n As Integer
    n = FreeFile
    Open "aliases.txt" For Input As #n
        Dim strHost As String, strAliases As String, strTemp As String
        Do While Not EOF(n)
            Line Input #n, strTemp
            strHost = Split(strTemp, "/")(0)
            strAliases = Split(strTemp, "/")(1)
            Dim strAlias() As String
            strAlias = Split(strAliases, ".")
            Dim i As Integer
            For i = LBound(strAlias) To UBound(strAlias)
                Call AddUser(strHost, strAlias(i))
            Next i
        Loop
    Close #n
End Sub

Public Sub AddUser(ByVal Nickname As String, ByVal UserHost As String)
    Debug.Print "AddUser(" & Nickname & ", " & UserHost & ")"
    On Error GoTo HostNotFound:
    Dim tempAlias As clsAliasData
    Set tempAlias = collAliases(UserHost)
    tempAlias.AddAlias Nickname
    'Call SaveData
    Exit Sub
HostNotFound:
    Debug.Print "AddUser.HostNotFound"
    Set tempAlias = New clsAliasData
    tempAlias.AddAlias Nickname
    tempAlias.Hostmask = UserHost
    collAliases.Add tempAlias, UserHost
    'Call SaveData
End Sub

Public Function GetAliases(ByVal UserHost As String)
    Debug.Print "GetAliases(" & UserHost & ")"
    On Error GoTo HostNotFound:
    Dim tempAlias As clsAliasData
    Set tempAlias = collAliases(UserHost)
    GetAliases = tempAlias.GetAliases
    Exit Function
HostNotFound:
    Debug.Print "GetAliases.HostNotFound"
End Function

Private Sub Class_Initialize()
    Set collAliases = New Collection
End Sub
