VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIRCPacketAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event OnChannelsFormed(ByVal NumberOfChannels As String)
Public Event OnGlobalUsers(ByVal Message As String)
Public Event OnLocalUsers(ByVal Message As String)
Public Event OnLoggedIn()
Public Event OnModeChange(ByVal ObjectAffected As String, ByVal Modes As String)
Public Event OnMOTDBegin(ByVal Message As String)
Public Event OnMOTDData(ByVal Message As String)
Public Event OnMOTDEnd(ByVal Message As String)
Public Event OnNameInChannel(ByVal Channel As String, ByVal Nickname As String)
Public Event OnNetworkInformation(ByVal Message As String)
Public Event OnNicknameChanged(ByVal OldNickname As String, ByVal NewNickname As String, ByVal UserHost As String)
Public Event OnNicknameInUse()
Public Event OnNoSuckNickChannel(ByVal UnknownData As String, ByVal Message As String)
Public Event OnOperatorsOnline(ByVal NumberOfOperators As String)
Public Event OnPrivateMessage(ByVal Destination As String, ByVal Nickname As String, ByVal UserHost As String, ByVal Message As String)
Public Event OnServerCreated(ByVal Message As String)
Public Event OnServerInfo(ByVal Message As String)
Public Event OnServerNotice(ByVal Destination As String, ByVal Message As String)
Public Event OnServerPing(ByVal PingData As String)
Public Event OnServerWelcome(ByVal Message As String)
Public Event OnSpecificServerInfo(ByVal Message As String)
Public Event OnTopicChange(ByVal Channel As String, ByVal Nickname As String, ByVal UserHost As String, ByVal Topic As String)
Public Event OnTopicIs(ByVal Channel As String, ByVal Topic As String)
Public Event OnTopicSetBy(ByVal Channel As String, ByVal Nickname As String, ByVal AtTime As String)
Public Event OnUserJoined(ByVal Channel As String, ByVal Nickname As String, ByVal UserHost As String)
Public Event OnUserLeft(ByVal Channel As String, ByVal Nickname As String, ByVal UserHost As String)
Public Event OnUserQuit(ByVal Nickname As String, ByVal UserHost As String, ByVal Message As String)
Public Event OnWhoisInfo(ByVal Nickname As String, ByVal Username As String, ByVal UserHost As String)

Public Sub ParseIRCPacket(ByRef IRCPacket As String)
    Dim IRCParams() As String
    Dim strSplit() As String
    Dim parseBase As Integer, i As Integer
    parseBase = 0
    
    IRCParams = Split(IRCPacket, ":")
    
    If Len(IRCParams(0)) = 0 Then
        parseBase = 1
        
        Dim commandParams() As String
        commandParams = Split(IRCParams(1), " ")
        Dim rest As String
        rest = Mid$(IRCPacket, InStr(2, IRCPacket, ":") + 1)
        
        If IsNumeric(commandParams(1)) Then
            Select Case commandParams(1)
                Case "001": ' Welcome
                    RaiseEvent OnServerWelcome(rest)
                Case "002": ' HostInfo
                    RaiseEvent OnServerInfo(rest)
                Case "003": ' ServerCreated
                    RaiseEvent OnServerCreated(rest)
                Case "004": ' Logged in
                    RaiseEvent OnLoggedIn
                Case "005": ' Server supports...
                    ' I hate this packet.
                Case "251": ' Netinfo
                    RaiseEvent OnNetworkInformation(rest)
                Case "252": ' Operators online
                    RaiseEvent OnOperatorsOnline(commandParams(3))
                Case "254": ' Channels formed
                    RaiseEvent OnChannelsFormed(commandParams(3))
                Case "255": ' Specific server info
                    RaiseEvent OnSpecificServerInfo(rest)
                Case "265": ' Local users
                    RaiseEvent OnLocalUsers(rest)
                Case "266": ' Global users
                    RaiseEvent OnGlobalUsers(rest)
                Case "311": ' Whois
                    RaiseEvent OnWhoisInfo(commandParams(3), commandParams(4), commandParams(5))
                Case "332": ' The topic is...
                    RaiseEvent OnTopicIs(commandParams(3), rest)
                Case "333": ' The topic was set by...
                    RaiseEvent OnTopicSetBy(commandParams(3), commandParams(4), commandParams(5))
                Case "353": ' Names list
                    strSplit = Split(rest, " ")
                    For i = LBound(strSplit) To UBound(strSplit)
                        RaiseEvent OnNameInChannel(commandParams(4), strSplit(i))
                    Next i
                Case "366": ' End of NAMES
                Case "372": ' MOTD Data
                    RaiseEvent OnMOTDData(rest)
                Case "375": ' Start of MOTD
                    RaiseEvent OnMOTDBegin(rest)
                Case "376": ' End of MOTD
                    RaiseEvent OnMOTDEnd(rest)
                Case "401": ' No such nick/channel
                    RaiseEvent OnNoSuckNickChannel(commandParams(3), rest)
                Case "433":
                    RaiseEvent OnNicknameInUse
                Case Else
                    Debug.Print "Unhandled numeric (" & commandParams(1) & "): " & IRCPacket
            End Select
        Else ' If Not IsNumeric
            Dim strTalker As String
        
            strSplit = Split(commandParams(0), "!")
            strTalker = strSplit(0)
            
            Select Case commandParams(1)
                Case "NOTICE"
                    ' Notice from the server
                    RaiseEvent OnServerNotice(commandParams(2), rest)
                Case "MODE"
                    ' Someone changed a mode.
                    RaiseEvent OnModeChange(commandParams(2), rest)
                Case "JOIN"
                    ' Someone joined.
                    RaiseEvent OnUserJoined(rest, strTalker, strSplit(1))
                Case "PART"
                    RaiseEvent OnUserLeft(rest, strTalker, strSplit(1))
                Case "NICK"
                    'RaiseEvent
                Case "PRIVMSG"
                    RaiseEvent OnPrivateMessage(commandParams(2), strTalker, strSplit(1), rest)
                Case "TOPIC"
                    RaiseEvent OnTopicChange(commandParams(2), strTalker, strSplit(1), rest)
                Case "QUIT"
                    RaiseEvent OnUserQuit(strTalker, strSplit(1), rest)
                Case Else
                    Debug.Print "Unhandled non-numeric (" & commandParams(1) & "): " & IRCPacket
            End Select
        End If
    Else
        parseBase = 0
        
        Dim cmpString As String
        cmpString = Trim(IRCParams(parseBase))
        
        Select Case cmpString
            Case "PING"
                RaiseEvent OnServerPing(IRCParams(parseBase + 1))
        End Select
    End If
End Sub
