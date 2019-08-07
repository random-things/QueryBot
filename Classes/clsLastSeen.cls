VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLastSeen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private collLastSeen As Collection

Public Sub SaveData()
    Dim n As Integer
    n = FreeFile
    Open "lastseen.txt" For Output As #n
        Dim tempLS As clsLastSeenData, tempString As String
        For Each tempLS In collLastSeen
            tempString = tempLS.Hostmask & "?" & tempLS.LastSeen & "?" & tempLS.LastSeenDoing & "?" & tempLS.Nickname
            Print #n, tempString
        Next
    Close #n
End Sub

Public Sub LoadData()
    Dim n As Integer
    n = FreeFile
    Open "lastseen.txt" For Input As #n
        Dim strHost As String, strLSD As String, strNick As String, dateDate As Date, strTemp As String, strStuff() As String
        Do While Not EOF(n)
            Line Input #n, strTemp
            strStuff = Split(strTemp, "?")
            strHost = strStuff(0)
            dateDate = strStuff(1)
            strLSD = strStuff(2)
            strNick = strStuff(3)
            
            Call AddLastSeenManual(strNick, strHost, strLSD, dateDate)
        Loop
    Close #n
End Sub

Public Sub AddLastSeenManual(ByVal Nickname As String, ByVal Hostmask As String, ByVal LastSeenDoing As String, ByVal LastSeen As Date)
    On Error GoTo NickNotFound:
    
    Dim tempLastSeen As clsLastSeenData
    Set tempLastSeen = collLastSeen(Nickname)
    
    tempLastSeen.Nickname = Nickname
    tempLastSeen.Hostmask = Hostmask
    tempLastSeen.LastSeenDoing = LastSeenDoing
    tempLastSeen.LastSeen = LastSeen
    
    Exit Sub

NickNotFound:
    Set tempLastSeen = New clsLastSeenData
    tempLastSeen.Nickname = Nickname
    tempLastSeen.Hostmask = Hostmask
    tempLastSeen.LastSeenDoing = LastSeenDoing
    tempLastSeen.LastSeen = LastSeen
    collLastSeen.Add tempLastSeen, Nickname
End Sub

Public Sub AddLastSeen(ByVal Nickname As String, ByVal Hostmask As String, ByVal LastSeenDoing As String)
    On Error GoTo NickNotFound:
    
    Dim tempLastSeen As clsLastSeenData
    Set tempLastSeen = collLastSeen(Nickname)
    
    tempLastSeen.Nickname = Nickname
    tempLastSeen.Hostmask = Hostmask
    tempLastSeen.LastSeenDoing = LastSeenDoing
    tempLastSeen.LastSeen = Now
    SaveData
    Exit Sub

NickNotFound:
    Set tempLastSeen = New clsLastSeenData
    tempLastSeen.Nickname = Nickname
    tempLastSeen.Hostmask = Hostmask
    tempLastSeen.LastSeenDoing = LastSeenDoing
    tempLastSeen.LastSeen = Now
    collLastSeen.Add tempLastSeen, Nickname
    SaveData
End Sub

Public Function GetLastSeen(ByVal Nickname As String) As clsLastSeenData
    On Error GoTo NickNotFound:

    Set GetLastSeen = collLastSeen(Nickname)

    Exit Function
    
NickNotFound:
    'Set GetLastSeen = Null
End Function

Private Sub Class_Initialize()
    Set collLastSeen = New Collection
End Sub
