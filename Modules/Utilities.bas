Attribute VB_Name = "Utilities"
Option Explicit

Public Function DebugOutput(ByVal sIn As String) As String
   Dim x1 As Long, y1 As Long
   Dim iLen As Long, iPos As Long
   Dim sB As String, sT As String
   Dim sOut As String
   Dim Offset As Long, sOffset As String
   
   iLen = Len(sIn)
   If iLen = 0 Then Exit Function
   sOut = ""
   Offset = 0
   For x1 = 0 To ((iLen - 1) \ 16)
       sOffset = Right$("0000" & Hex(Offset), 4)
       sB = String(48, " ")
       sT = "................"
       For y1 = 1 To 16
           iPos = 16 * x1 + y1
           If iPos > iLen Then Exit For
           Mid(sB, 3 * (y1 - 1) + 1, 2) = Right("00" & Hex(Asc(Mid(sIn, iPos, 1))), 2) & " "
           Select Case Asc(Mid(sIn, iPos, 1))
           Case 0, 9, 10, 13
           Case Else
               Mid(sT, y1, 1) = Mid(sIn, iPos, 1)
           End Select
       Next y1
       If Len(sOut) > 0 Then sOut = sOut & vbCrLf
       sOut = sOut & sOffset & ":  "
       sOut = sOut & sB & "  " & sT
       Offset = Offset + 16
   Next x1
   DebugOutput = sOut
End Function

