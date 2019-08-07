VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Software related variables
Private WithEvents Parser As clsIRCPacketAnalyzer
Attribute Parser.VB_VarHelpID = -1
Private WithEvents sckIRC As MSWinsockLib.Winsock
Attribute sckIRC.VB_VarHelpID = -1

Private bufIRC As clsIRCBuffer

' Protocol related variables
Private strServer As String
Private strPort As String
Private strNick As String
Private strUser As String
Private strHost As String
Private strName As String

' Status related variables
Private bConnecting As Boolean
Private bConnected As Boolean

' Protocol related constants
Private Const DEFAULT_SERVER As String = "irc.somewhere.org"
Private Const DEFAULT_PORT As String = "6667"
Private Const DEFAULT_NICK As String = "Bot"
Private Const DEFAULT_USER As String = "Bot"
Private Const DEFAULT_HOST As String = "vb.bot"
Private Const DEFAULT_NAME As String = "VB Bot"

' Socket Events
Public Event OnConnecting(ByVal Server As String, ByVal Port As String)
Public Event OnConnected(ByVal Server As String, ByVal ServerIP As String, ByVal Port As String)
Public Event OnDisconnected()
Public Event OnError(ByVal Number As Integer, ByVal Description As String)

' IRC Events
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

Public Property Let Socket(ByRef IRCSocket As MSWinsockLib.Winsock)
    Set sckIRC = IRCSocket
End Property

Public Property Let Server(ByVal NewServer As String)
    strServer = NewServer
End Property
Public Property Get Server() As String
    Server = strServer
End Property
Public Property Let Port(ByVal NewPort As String)
    strPort = NewPort
End Property
Public Property Get Port() As String
    Port = strPort
End Property
Public Property Let Nick(ByVal NewNick As String)
    strNick = NewNick
End Property
Public Property Get Nick() As String
    Nick = strNick
End Property
Public Property Let User(ByVal NewUser As String)
    strUser = NewUser
End Property
Public Property Get User() As String
    User = strUser
End Property
Public Property Let Host(ByVal NewHost As String)
    strHost = NewHost
End Property
Public Property Get Host() As String
    Host = strHost
End Property
Public Property Let Name(ByVal NewName As String)
    strName = NewName
End Property
Public Property Get Name() As String
    Name = strName
End Property

Private Sub EnsureDefaults()
    If StrComp(strServer, vbNullString) = 0 Then
        strServer = DEFAULT_SERVER
    End If
    
    If StrComp(strPort, vbNullString) = 0 Then
       strPort = DEFAULT_PORT
    End If
    
    If StrComp(strNick, vbNullString) = 0 Then
        strNick = DEFAULT_NICK
    End If
    
    If StrComp(strUser, vbNullString) = 0 Then
        strUser = DEFAULT_USER
    End If
    
    If StrComp(strHost, vbNullString) = 0 Then
        strHost = DEFAULT_HOST
    End If
    
    If StrComp(strName, vbNullString) = 0 Then
        strName = DEFAULT_NAME
    End If
End Sub

Public Function IRCConnect() As Boolean
    ' Make sure that relevant data is filled out.
    Call EnsureDefaults
    
    bConnecting = True
    RaiseEvent OnConnecting(strServer, strPort)
    sckIRC.Connect strServer, strPort
End Function

Public Sub Send(ByVal strOut As String)
    Call IRCWrite(strOut)
End Sub

Private Function IRCWrite(ByVal strOut As String) As Boolean
    If bConnected = True Then
        sckIRC.SendData strOut & vbCrLf
    End If
End Function

Public Sub JoinChannel(ByVal Channel As String)
    Call IRCWrite("JOIN " & Channel)
End Sub

Public Sub LeaveChannel(ByVal Channel As String)
    Call IRCWrite("PART " & Channel)
End Sub

Public Sub PrivateMessage(ByVal Destination As String, ByVal Message As String)
    Call IRCWrite("PRIVMSG " & Destination & " :" & Message)
End Sub

Private Sub Class_Initialize()
    Set Parser = New clsIRCPacketAnalyzer
    Set bufIRC = New clsIRCBuffer
End Sub

Private Sub Parser_OnChannelsFormed(ByVal NumberOfChannels As String)
    RaiseEvent OnChannelsFormed(NumberOfChannels)
End Sub

Private Sub Parser_OnGlobalUsers(ByVal Message As String)
    RaiseEvent OnGlobalUsers(Message)
End Sub

Private Sub Parser_OnLocalUsers(ByVal Message As String)
    RaiseEvent OnLocalUsers(Message)
End Sub

Private Sub Parser_OnLoggedIn()
    RaiseEvent OnLoggedIn
End Sub

Private Sub Parser_OnModeChange(ByVal ObjectAffected As String, ByVal Modes As String)
    RaiseEvent OnModeChange(ObjectAffected, Modes)
End Sub

Private Sub Parser_OnMOTDBegin(ByVal Message As String)
    RaiseEvent OnMOTDBegin(Message)
End Sub

Private Sub Parser_OnMOTDData(ByVal Message As String)
    RaiseEvent OnMOTDData(Message)
End Sub

Private Sub Parser_OnMOTDEnd(ByVal Message As String)
    RaiseEvent OnMOTDEnd(Message)
End Sub

Private Sub Parser_OnNameInChannel(ByVal Channel As String, ByVal Nickname As String)
    RaiseEvent OnNameInChannel(Channel, Nickname)
End Sub

Private Sub Parser_OnNetworkInformation(ByVal Message As String)
    RaiseEvent OnNetworkInformation(Message)
End Sub

Private Sub Parser_OnNicknameChanged(ByVal OldNickname As String, ByVal NewNickname As String, ByVal UserHost As String)
    RaiseEvent OnNicknameChanged(OldNickname, NewNickname, UserHost)
End Sub

Private Sub Parser_OnNicknameInUse()
    RaiseEvent OnNicknameInUse
End Sub

Private Sub Parser_OnNoSuckNickChannel(ByVal UnknownData As String, ByVal Message As String)
    RaiseEvent OnNoSuckNickChannel(UnknownData, Message)
End Sub

Private Sub Parser_OnOperatorsOnline(ByVal NumberOfOperators As String)
    RaiseEvent OnOperatorsOnline(NumberOfOperators)
End Sub

Private Sub Parser_OnPrivateMessage(ByVal Destination As String, ByVal Nickname As String, ByVal UserHost As String, ByVal Message As String)
    RaiseEvent OnPrivateMessage(Destination, Nickname, UserHost, Message)
End Sub

Private Sub Parser_OnServerCreated(ByVal Message As String)
    RaiseEvent OnServerCreated(Message)
End Sub

Private Sub Parser_OnServerInfo(ByVal Message As String)
    RaiseEvent OnServerInfo(Message)
End Sub

Private Sub Parser_OnServerNotice(ByVal Destination As String, ByVal Message As String)
    RaiseEvent OnServerNotice(Destination, Message)
End Sub

Private Sub Parser_OnServerPing(ByVal PingData As String)
    Debug.Print "Got pinged."
    Call IRCWrite("PONG " & PingData)
End Sub

Private Sub Parser_OnServerWelcome(ByVal Message As String)
    RaiseEvent OnServerWelcome(Message)
End Sub

Private Sub Parser_OnSpecificServerInfo(ByVal Message As String)
    RaiseEvent OnSpecificServerInfo(Message)
End Sub

Private Sub Parser_OnTopicChange(ByVal Channel As String, ByVal Nickname As String, ByVal UserHost As String, ByVal Topic As String)
    RaiseEvent OnTopicChange(Channel, Nickname, UserHost, Topic)
End Sub

Private Sub Parser_OnTopicIs(ByVal Channel As String, ByVal Topic As String)
    RaiseEvent OnTopicIs(Channel, Topic)
End Sub

Private Sub Parser_OnTopicSetBy(ByVal Channel As String, ByVal Nickname As String, ByVal AtTime As String)
    RaiseEvent OnTopicSetBy(Channel, Nickname, AtTime)
End Sub

Private Sub Parser_OnUserJoined(ByVal Channel As String, ByVal Nickname As String, ByVal UserHost As String)
    RaiseEvent OnUserJoined(Channel, Nickname, UserHost)
End Sub

Private Sub Parser_OnUserLeft(ByVal Channel As String, ByVal Nickname As String, ByVal UserHost As String)
    RaiseEvent OnUserLeft(Channel, Nickname, UserHost)
End Sub

Private Sub Parser_OnUserQuit(ByVal Nickname As String, ByVal UserHost As String, ByVal Message As String)
    RaiseEvent OnUserQuit(Nickname, UserHost, Message)
End Sub

Private Sub Parser_OnWhoisInfo(ByVal Nickname As String, ByVal Username As String, ByVal UserHost As String)
    RaiseEvent OnWhoisInfo(Nickname, Username, UserHost)
End Sub

Private Sub sckIRC_Close()
    bConnected = False
    
    RaiseEvent OnDisconnected
End Sub

Private Sub sckIRC_Connect()
    ' Update status
    bConnecting = False
    bConnected = True
    
    RaiseEvent OnConnected(sckIRC.RemoteHost, sckIRC.RemoteHostIP, sckIRC.RemotePort)
    
    Call IRCWrite("NICK " & strNick)
    Call IRCWrite("USER " & strUser & " " & strHost & " " & strServer & " :" & strName)
End Sub

Private Sub sckIRC_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    
    sckIRC.GetData strData, vbString
    
    bufIRC.AddData strData
    
    While (bufIRC.IsFullPacket = True)
        Call Parser.ParseIRCPacket(bufIRC.GetPacket)
    Wend
End Sub

Private Sub sckIRC_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent OnDisconnected
    RaiseEvent OnError(Number, Description)
    sckIRC.Close
End Sub
