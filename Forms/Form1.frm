VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents IRC As clsIRC
Attribute IRC.VB_VarHelpID = -1
Private QA As clsQA
Private LS As clsLastSeen
Private Alias As clsAliases

Private lastChannel As String
Private lastUser As String
Private lastParam As String
Private lastCommand As String

Private Sub Form_Load()
    Set IRC = New clsIRC
    Set QA = New clsQA
    Set LS = New clsLastSeen
    Set Alias = New clsAliases
    IRC.Socket = Winsock1
    
    IRC.Host = "my.host"
    IRC.Name = "Query Bot v1.0"
    IRC.Nick = "Bot"
    IRC.Port = "6667"
    IRC.Server = "192.168.0.10"
    IRC.User = "QueryBot"
    
    IRC.IRCConnect
End Sub

Private Sub IRC_OnConnected(ByVal Server As String, ByVal ServerIP As String, ByVal Port As String)
    Debug.Print "Connected to: " & Server & "(" & ServerIP & "):" & Port
End Sub

Private Sub IRC_OnConnecting(ByVal Server As String, ByVal Port As String)
    Debug.Print "Connecting to: " & Server & ":" & Port
End Sub

Private Sub IRC_OnNameInChannel(ByVal Channel As String, ByVal Nickname As String)
    Debug.Print Channel & "|" & Nickname & " is here."
End Sub

Private Sub IRC_OnPrivateMessage(ByVal Destination As String, ByVal Nickname As String, ByVal UserHost As String, ByVal Message As String)
    UserHost = Split(UserHost, "@")(1)
    Debug.Print Destination & "|" & Nickname & "> " & Message
    
    LS.AddLastSeen Nickname, UserHost, "talking"
    
    If Nickname = "a" Then
        If Left$(Message, 1) = "." Then
            Dim strCommands() As String
            strCommands = Split(Message, " ")
            strCommands(0) = Mid$(strCommands(0), 2)
            
            Dim tempMessage As String
            tempMessage = Mid$(Message, Len(strCommands(0)) + 3)
            'Debug.Print "Message -> " & Message
            
            Select Case strCommands(0)
                Case "seen"
                    If strCommands(1) = Nickname Then
                        IRC.PrivateMessage Destination, "You are already here!"
                        Exit Sub
                    End If
                
                    Dim lsd As clsLastSeenData
                    Set lsd = LS.GetLastSeen(strCommands(1))
                    
                    If Not (lsd Is Nothing) Then
                        Dim lstime As Long
                        lstime = DateDiff("s", Now, lsd.LastSeen)
                        IRC.PrivateMessage Destination, strCommands(1) & " was last seen " & lsd.LastSeenDoing & " " & ParseSeconds(DateDiff("s", lsd.LastSeen, Now)) & " ago on " & FormatDateTime(lsd.LastSeen, vbLongDate) & " at " & FormatDateTime(lsd.LastSeen, vbLongTime) & "."
                    Else
                        Set lsd = LS.GetLastSeen(Mid$(Destination, 2) & strCommands(1))
                        If Not (lsd Is Nothing) Then
                            lstime = DateDiff("s", Now, lsd.LastSeen)
                            IRC.PrivateMessage Destination, Mid$(Destination, 2) & strCommands(1) & " was last seen " & lsd.LastSeenDoing & " " & ParseSeconds(DateDiff("s", lsd.LastSeen, Now)) & " ago on " & FormatDateTime(lsd.LastSeen, vbLongDate) & " at " & FormatDateTime(lsd.LastSeen, vbLongTime) & "."
                        Else
                            IRC.PrivateMessage Destination, "The user (" & strCommands(1) & ") was not found."
                        End If
                    End If
                Case "join"
                    IRC.JoinChannel strCommands(1)
                Case "part"
                    IRC.LeaveChannel strCommands(1)
                Case "say"
                    'IRC.PrivateMessage strCommands(1), Mid$(Message, 6 + Len(strCommands(1)) + 1)
                    IRC.PrivateMessage strCommands(1), Mid$(tempMessage, Len(strCommands(1)) + 2)
                Case "qa"
                    Dim strQ As String, strA As String
                    strQ = Split(tempMessage, "|")(0)
                    strA = Split(tempMessage, "|")(1)
                    QA.AddQAPair strQ, strA
                    Exit Sub
                Case "match"
                    IRC.PrivateMessage Destination, QA.MatchQA(tempMessage)
                Case "remove"
                    QA.RemoveQA tempMessage
                Case "save"
                    QA.SaveData
                    LS.SaveData
                    Alias.SaveData
                Case "load"
                    QA.LoadData
                    LS.LoadData
                    Alias.LoadData
                Case "whois"
                    lastChannel = Destination
                    lastUser = tempMessage
                    lastCommand = "whois"
                    Call IRC.Send("WHOIS " & tempMessage)
                Case "ip"
                    lastChannel = Destination
                    lastUser = tempMessage
                    lastCommand = "ip"
                    Call IRC.Send("WHOIS " & tempMessage)
            End Select
        End If
    End If
    
    If Left$(Message, 1) = "." Then

        strCommands = Split(Message, " ")
        strCommands(0) = Mid$(strCommands(0), 2)
        
        tempMessage = Mid$(Message, Len(strCommands(0)) + 3)
        'Debug.Print "Message2 -> " & tempMessage
        
        Dim tempString As String
        Select Case strCommands(0)
            Case "freesay"
                'IRC.PrivateMessage strCommands(1), Mid$(Message, 6 + Len(strCommands(1)) + 1)
                IRC.PrivateMessage Destination, Nickname & " says: " & tempMessage
            Case "serve"
                tempString = Split(tempMessage, " ")(0)
                IRC.PrivateMessage Destination, IRC.Nick & " serves " & tempString & " a " & Mid$(tempMessage, Len(tempString) + 2) & "."
            Case "topic"
                IRC.Send "TOPIC " & Destination & " :" & tempMessage
            Case "names"
                lastChannel = Destination
                lastUser = tempMessage
                lastCommand = "names"
                lastParam = Nickname
                Call IRC.Send("WHOIS " & tempMessage)
            Case "namecount"
                lastChannel = Destination
                lastUser = tempMessage
                lastCommand = "namecount"
                lastParam = Nickname
                Call IRC.Send("WHOIS " & tempMessage)
        End Select
    End If
    
    Dim retString As String
    retString = QA.CheckQA(Message)
    If Len(retString) > 0 Then
        retString = Replace(retString, "$name$", Nickname)
        retString = Replace(retString, "$host$", UserHost)
        
        IRC.PrivateMessage Destination, retString
    End If
End Sub

Private Sub IRC_OnUserJoined(ByVal Channel As String, ByVal Nickname As String, ByVal UserHost As String)
    UserHost = Split(UserHost, "@")(1)
    Debug.Print Channel & "|" & Nickname & " joined."
    
    LS.AddLastSeen Nickname, UserHost, "joining"
    Alias.AddUser Nickname, UserHost
End Sub

Private Sub IRC_OnUserLeft(ByVal Channel As String, ByVal Nickname As String, ByVal UserHost As String)
    UserHost = Split(UserHost, "@")(1)
    Debug.Print Channel & "|" & Nickname & " left."
    
    LS.AddLastSeen Nickname, UserHost, "leaving"
    Alias.AddUser Nickname, UserHost
End Sub

Private Sub IRC_OnUserQuit(ByVal Nickname As String, ByVal UserHost As String, ByVal Message As String)
    UserHost = Split(UserHost, "@")(1)
    Debug.Print Nickname & " quit: " & Message
    
    LS.AddLastSeen Nickname, UserHost, "quitting"
    Alias.AddUser Nickname, UserHost
End Sub

Private Sub IRC_OnWhoisInfo(ByVal Nickname As String, ByVal Username As String, ByVal UserHost As String)
    Select Case lastCommand
        Case "ip":
            IRC.PrivateMessage lastChannel, lastUser & "'s IP: " & UserHost
        Case "whois":
            IRC.PrivateMessage lastChannel, lastUser & " has nickname " & Nickname & " - " & lastUser & " has username " & Username & " - and " & lastUser & " is connecting from: " & UserHost
        Case "names":
            Dim retString As String
            retString = Alias.GetAliases(UserHost)
            If Len(retString) > 0 Then
                IRC.PrivateMessage lastChannel, lastUser & " has used the names: " & Replace(Alias.GetAliases(UserHost), ".", ", ")
            Else
                IRC.PrivateMessage lastChannel, "The user " & lastUser & " was not found."
            End If
        Case "namecount":
            retString = Alias.GetAliases(UserHost)
            If Len(retString) > 0 Then
                'IRC.PrivateMessage lastChannel, lastUser & " has used the names: " & Replace(Alias.GetAliases(UserHost), ".", ", ")
                Dim tempString() As String
                tempString = Split(retString, ".")
                IRC.PrivateMessage lastChannel, lastUser & " has used a total of " & UBound(tempString) + 1 & IIf(UBound(tempString) > 0, " names.", " name.")
            Else
                IRC.PrivateMessage lastChannel, "The user " & lastUser & " was not found."
            End If
    End Select
End Sub
