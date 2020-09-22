Attribute VB_Name = "Battlenet2"
 Public Sub OpenFile(File As String)
 Dim MyString As String
    On Error Resume Next
    frmEdit.TextBox.text = ""
    Open App.Path & "\" & File For Input As #1


    While Not EOF(1)
        Input #1, MyString$


        DoEvents
            frmEdit.TextBox = frmEdit.TextBox & MyString$ & vbCrLf
        Wend
        Close #1
        Close #1
        Close #1
    End Sub

Sub Pause(hInterval As Long)
    Dim hCurrent As Long
    hInterval = hInterval * 100
    hCurrent = GetTickCount
    Do While GetTickCount - hCurrent < Val(hInterval)
        DoEvents
    Loop
End Sub


Function GetStuff(appname As String, key As String) As String
Dim sFile As String
Dim sDefault As String
Dim lSize As Integer
Dim L As Long
Dim sUser As String
sUser = Space$(128)
lSize = Len(sUser)
sFile = App.Path & "\Config.ini"
sDefault = ""
L = GetPrivateProfileString(appname, key, sDefault, sUser, lSize, sFile)
sUser = Mid(sUser, 1, InStr(sUser, Chr(0)) - 1)
GetStuff = sUser
End Function
Public Sub SendChat(text As String)
'this sub sends data to battle.net and adds it to the chat
If frmBot.Winsock.State = sckConnected Then
frmBot.Winsock.SendData text & vbCrLf
If Left(text, 1) = "/" Then Exit Sub
AddChat "<" & frmBot.strName & ">", &HFFFF00, " " & text & vbCrLf, vbWhite
Else
MsgBox "You Are Not Connected!", vbExclamation, "Error"
End If
End Sub
Public Sub Send(Data)
'this sub sends data to battle.net without adding to the chat
If frmBot.Winsock.State = sckConnected Then
frmBot.Winsock.SendData text & vbCrLf
Else
MsgBox "You Are Not Connected!", vbExclamation, "Error"
End If
End Sub
Public Function GetIconCode(Client As String, Optional Flags As String) As Integer
'this sub takes the client info recieved by battle.net and
'converts it into an integer, the integer represents an image
'in the image list on frmbot, that image is placed in the list
Dim Code As Integer
Select Case Client
  Case "[CHAT]": Code = ICON_CHAT
  Case "[DRTL]": Code = ICON_DRTL
  Case "[DSHR]": Code = ICON_DSHR
  Case "[D2XP]": Code = ICON_D2XP
  Case "[STAR]": Code = ICON_STAR
  Case "[SEXP]": Code = ICON_SEXP
  Case "[SSHR]": Code = ICON_SSHR
  Case "[W2BN]": Code = ICON_W2BN
  Case "[D2DV]": Code = ICON_D2DV
  Case "[JSTR]": Code = ICON_JSTR
  Case "[BRX]":  Code = ICON_IGNORE
  Case Else: Code = ICON_QUES
End Select
If Mid(Flags, 4, 1) = "2" Then Code = ICON_GAVEL
If Mid(Flags, 3, 2) = "20" Then Code = ICON_IGNORE
If Mid(Flags, 3, 2) = "30" Then Code = ICON_IGNORE
If Mid(Flags, 4, 1) = "1" Then Code = ICON_BLIZZ
If Mid(Flags, 4, 1) = "4" Then Code = ICON_MEGA
GetIconCode = Code
End Function
Public Function ChangeFlags(Flags As String)
'this isnt really used anymore by chatbots since the no operators
'patch, but when someone gets ops or something else changes
'this is the sub
Dim Code As Integer
If Mid(Flags, 4, 1) = "2" Then
Code = ICON_GAVEL
ElseIf Mid(Flags, 3, 2) = "20" Then
Code = ICON_IGNORE
ElseIf Mid(Flags, 3, 2) = "30" Then
Code = ICON_IGNORE
ElseIf Mid(Flags, 4, 1) = "1" Then
Code = ICON_BLIZZ
ElseIf Mid(Flags, 4, 1) = "4" Then
Code = ICON_MEGA
Else
Code = ICON_QUES
End If
ChangeFlags = Code
End Function
Public Sub AddChat(ByVal txtOne As String, ByVal clrOne As Long, Optional ByVal txtTwo As String, Optional ByVal clrTwo As Long, Optional strType As String)
'this adds COLORED text to the chat, 2 colors in this example
With frmBot.txtChannel
  .SelStart = Len(.text)
  .SelLength = 0
  .SelColor = clrOne
  .SelText = txtOne
End With

If Len(txtTwo) > 0 Then
    With frmBot.txtChannel
    .SelStart = Len(.text)
    .SelLength = 0
    .SelColor = clrTwo
    .SelText = txtTwo
    End With
End If
End Sub
Public Function ProfanityUnfilter(strmessage As String) As String
'this searches for the filtered message and replaces it with the actual word.
If InStr(strmessage, "!&$%@") <> 0 Then strmessage = Replace(strmessage, "!&$%@", "pussy")
If InStr(strmessage, "$%@%") <> 0 Then strmessage = Replace(strmessage, "$%@%", "clit")
If InStr(strmessage, "@$%!@%&") <> 0 Then strmessage = Replace(strmessage, "@$%!@%&", "asshole")
If InStr(strmessage, "#@%$!") <> 0 Then strmessage = Replace(strmessage, "#@%$!", "bitch")
If InStr(strmessage, "$!@!$") <> 0 Then strmessage = Replace(strmessage, "$!@!$", "chink")
If InStr(strmessage, "$!$%") <> 0 Then strmessage = Replace(strmessage, "$!$%", "cock")
If InStr(strmessage, "$&!%") <> 0 Then strmessage = Replace(strmessage, "$&!%", "cunt")
If InStr(strmessage, "%@$%") <> 0 Then strmessage = Replace(strmessage, "%@$%", "dick")
If InStr(strmessage, "%@%&!") <> 0 Then strmessage = Replace(strmessage, "%@%&!", "dildo")
If InStr(strmessage, "&#&$%") <> 0 Then strmessage = Replace(strmessage, "&#&$%", "erect")
If InStr(strmessage, "!@!@!%") <> 0 Then strmessage = Replace(strmessage, "!@!@!%", "faggot")
If InStr(strmessage, "!&$%") <> 0 Then strmessage = Replace(strmessage, "!&$%", "fuck")
If InStr(strmessage, "!@!$") <> 0 Then strmessage = Replace(strmessage, "!@!$", "gook")
If InStr(strmessage, "$@$&") <> 0 Then strmessage = Replace(strmessage, "$@$&", "kike")
If InStr(strmessage, "$%$") <> 0 Then strmessage = Replace(strmessage, "$%$", "kkk")
If InStr(strmessage, "$%&!") <> 0 Then strmessage = Replace(strmessage, "$%&!", "klux")
If InStr(strmessage, "%&$#@#!") <> 0 Then strmessage = Replace(strmessage, "%&$#@#!", "lesbian")
If InStr(strmessage, "&@$%&#$@&") <> 0 Then strmessage = Replace(strmessage, "&@$%&#$@&", "masterbate")
If InStr(strmessage, "!@!@#") <> 0 Then strmessage = Replace(strmessage, "!@!@#", "nigga")
If InStr(strmessage, "!@!@&#") <> 0 Then strmessage = Replace(strmessage, "!@!@&#", "nigger")
If InStr(strmessage, "!@!@%&") <> 0 Then strmessage = Replace(strmessage, "!@!@%&", "nipple")
If InStr(strmessage, "!#!@$&") <> 0 Then strmessage = Replace(strmessage, "!#!@$&", "orgasm")
If InStr(strmessage, "!&!@$") <> 0 Then strmessage = Replace(strmessage, "!&!@$", "penis")
If InStr(strmessage, "$!@%") <> 0 Then strmessage = Replace(strmessage, "$!@%", "shit")
If InStr(strmessage, "!@!#&") <> 0 Then strmessage = Replace(strmessage, "!@!#&", "whore")
If InStr(strmessage, "$%&%") <> 0 Then strmessage = Replace(strmessage, "$%&%", "slut")
If InStr(strmessage, "!@!@!@") <> 0 Then strmessage = Replace(strmessage, "!@!@!@", "vagina")
ProfanityUnfilter = strmessage
End Function
