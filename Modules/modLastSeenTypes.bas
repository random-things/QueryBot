Attribute VB_Name = "modLastSeenFunctions"
Option Explicit

Public Function ParseSeconds(ByVal numSeconds As Long) As String
    'Debug.Print "ParseSeconds -> " & numSeconds
    Dim Minutes As Long
    Minutes = numSeconds / 60
    
    Dim Hours As Long
    Hours = Minutes / 60
    
    Dim Days As Long
    Days = Hours / 24

    Hours = Hours Mod 24
    If Days > 12 Then Days = Days - 1
    Minutes = Minutes Mod 60
    If Minutes > 30 Then Hours = Hours - 1
    numSeconds = numSeconds Mod 60
    If numSeconds > 30 Then Minutes = Minutes - 1
    
    Dim numCommas As Long
    numCommas = 0
    If Days > 0 Then numCommas = numCommas + 1
    If Hours > 0 Then numCommas = numCommas + 1
    If Minutes > 0 Then numCommas = numCommas + 1
    
    If Days > 0 Then
        ParseSeconds = CStr(Days) & IIf(Days = 1, " day", " days")
        If numCommas > 1 Then ParseSeconds = ParseSeconds & ","
    End If
    
    If Hours > 0 Then
        ParseSeconds = ParseSeconds & " " & CStr(Hours) & IIf(Hours = 1, " hour", " hours")
        If numCommas > 1 Then ParseSeconds = ParseSeconds & ","
    End If
    
    If Minutes > 0 Then
        ParseSeconds = ParseSeconds & " " & CStr(Minutes) & IIf(Minutes = 1, " minute", " minutes")
        If numCommas > 1 Then ParseSeconds = ParseSeconds & ","
    End If
    
    If numSeconds > 0 Then
        If numCommas > 0 Then ParseSeconds = ParseSeconds & " and"
        ParseSeconds = ParseSeconds & " " & CStr(numSeconds) & IIf(numSeconds = 1, " second", " seconds")
    End If
    
    If Right$(ParseSeconds, 1) = "," Then ParseSeconds = Left$(ParseSeconds, Len(ParseSeconds) - 1)
    
    ParseSeconds = LTrim(ParseSeconds)
End Function
