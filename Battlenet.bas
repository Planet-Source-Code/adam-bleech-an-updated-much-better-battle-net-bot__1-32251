Attribute VB_Name = "Battlenet"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Battle.net String Parser                           '''
''' Completley rebuilt by Adam Bleech (Atom)           '''
''' Origninal by Michael Horbatsch (Maglor)            '''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
DefInt I
DefStr S
DefLng L
Declare Function GetPrivateProfileSection Lib "KERNEL32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Const ICON_GAVEL    As Integer = 1 'Operator
Public Const ICON_CHAT     As Integer = 2 'Chat Bot
Public Const ICON_DRTL     As Integer = 3 'Diablo
Public Const ICON_DSHR     As Integer = 4 'Diablo Shareware
Public Const ICON_D2DV     As Integer = 5 'Diablo 2
Public Const ICON_STAR     As Integer = 6 'Starcraft
Public Const ICON_SSHR     As Integer = 7 'Starcraft Shareware
Public Const ICON_SPAWN    As Integer = 8 'Starcraft Spawn [NuLL]
Public Const ICON_SEXP     As Integer = 9 'Starcraft Brood war
Public Const ICON_W2BN     As Integer = 10 'Warcraft 2
Public Const ICON_JSTR     As Integer = 11 'Starcraft Japan
Public Const ICON_QUES     As Integer = 12 'Unknown Game
Public Const ICON_D2XP     As Integer = 13 'Diablo 2 Lord of Destruction
Public Const ICON_IGNORE   As Integer = 14 'BRX (ignored)
Public Const ICON_BLIZZ    As Integer = 15 'Blizzard Rep
Public Const ICON_BNET     As Integer = 16 'B.net Icon
Public Const ICON_MEGA     As Integer = 17 'Megaphone Icon
Public Const ICON_SHADES   As Integer = 18 'Battle.net Special Guest
Public IniUser As String, IniPass As String, IniServer As String, Database As Variant, TheChannel As String, IdleTick As Long 'declares variables for loading ini and misc use













Public Sub Connect(ctlSock As Control, Optional strPort = 6112)
    
    'this part loads the Config.Ini
    IniUser = GetStuff("Login", "User")
    IniPass = GetStuff("Login", "Pass")
    IniServer = GetStuff("Login", "Server")
    AddChat "Loaded Config.ini" & vbCrLf, vbGreen
    
    
    'this part connects to a Battle.net server
    
    'ctlSock  : A Microsoft Winsock Control v5.0+
    'strServer: The Battle.net server
    'strPort  : The remote port (don't change this unless you know what you're doing)
    If ctlSock.State <> 0 Then ctlSock.Disconnect
    With ctlSock
        .RemoteHost = IniServer
        .RemotePort = strPort
        .Protocol = sckTCPProtocol
        .Connect
    End With
End Sub

Public Sub Login(ctlSock As Control, strUser, strPass)
    'Login to the Battle.net Server. This should be called once a
    'connection has been established. The event Winsock_Connect
    'indicates that a connection has been established.
    
    'ctlSock    : The Winsock Control which established the connection
    'strUser    : Your username (use GUEST or ANONYMOUS if you don't have one)
    'strPassword: Your password (GUEST and ANONYMOUS don't require passwords)

    If ctlSock.State <> 0 Then
        ctlSock.SendData Chr(3) & Chr(4) & strUser & vbCrLf & strPass & vbCrLf
    End If
End Sub

Public Sub Disconnect(ctlSock As Control)
    'Disconnect from Battle.net
    
    'ctlSock: The Winsock Control
    ctlSock.Close
End Sub

Public Sub ParseData(strData)
    'Parse a data string which has been received from Battle.net.
    'This should be called whenever a Winsock_DataArrival event
    'occurs, where strData is the received data string.
    
    Dim strStrings() As String, i
    
    'Sometimes Battle.net puts several strings into one line,
    'separated by CrLf's. Here we split and parse them.
    
    SplitStr strData, strStrings, vbCrLf
    For i = 1 To UBound(strStrings)
        Parse2 strStrings(i)
    Next i
End Sub

Private Sub Parse2(ByVal strData As String)
On Error Resume Next
    'This subroutine does the actual parsing. It gets called
    'by the ParseData subroutine. There is no need to call it
    'from outside.
    
    Dim strData1, strData2, i
    Dim strArgs() As String
    
    'Sometimes, Battle.net puts a LineFeed character at the
    'beginning of strings. We need to trim it.
    If Left(strData, 1) = Chr(10) Then
        strData = Right(strData, Len(strData) - 1)
    End If
    
    'Check whether this is normal text, or an event
    If Val(Left(strData, 4)) < 1001 Or Val(Left(strData, 4)) > 3000 Then
        Event_Unknown strData
        Exit Sub
    End If
    
    'If this is an event, then we separate the arguments, and quoted text
    i = InStr(1, strData, Chr(34), vbBinaryCompare)
    If i <> 0 Then
        strData1 = Left(strData, i - 2)
        strData2 = Mid(strData, i + 1, Len(strData) - i - 1)
    Else
        strData1 = strData
        strData2 = ""
    End If
    SplitStr strData1, strArgs(), " "
    
    'Call Appropriate Event based on ID
    Select Case Val(strArgs(1))
        Case 1001: Event_User strArgs(3), strArgs(4), strArgs(5)
        Case 1002: Event_Join strArgs(3), strArgs(4), strArgs(5)
        Case 1003: Event_Leave strArgs(3), strArgs(4)
        Case 1004: Event_RecvWhisper strArgs(3), strArgs(4), strData2
        Case 1005: Event_Talk strArgs(3), strArgs(4), strData2
        Case 1006: Event_Broadcast strData2
        Case 1007: Event_Channel strData2
        Case 1009: Event_Flags strArgs(3), strArgs(4)
        Case 1010: Event_SendWhisper strArgs(3), strArgs(4), strData2
        Case 1013: Event_ChannelFull strData2
        Case 1014: Event_ChannelNotExist strData2
        Case 1015: Event_ChannelRestricted strData2
        Case 1016: Event_Info strData2
        Case 1018: Event_Info strData2
        Case 1019: Event_Error strData2
        Case 1023: Event_Emote strArgs(3), strArgs(4), strData2
        Case 2010: Event_Name strArgs(3)
        Case 3000: Event_Info strData2
    End Select
End Sub

Public Function SplitStr(ByVal OriginalString As String, ByRef ReturnArray() As String, ByVal Delimiter As String) As Long
    'Used to split a string by delimiters, into a dynamic array, and returns
    'the ubound of the array, i use this since vb5 doesnt have the split() function
    
    Dim sItem, lArrCnt
    Dim lLen, lPos
    
    lArrCnt = 0
    lLen = Len(OriginalString)

    Do
        lPos = InStr(1, OriginalString, Delimiter, vbTextCompare)
        If lPos <> 0 Then
            sItem = Left$(OriginalString, lPos - 1)
            OriginalString = Mid$(OriginalString, lPos + 1)
            If sItem <> "" Then
                lArrCnt = lArrCnt + 1
                If lArrCnt = 1 Then
                    ReDim ReturnArray(1 To lArrCnt) As String
                Else
                    ReDim Preserve ReturnArray(1 To lArrCnt) As String
                End If
                ReturnArray(lArrCnt) = sItem
            End If
        End If
    Loop While lPos <> 0

    If OriginalString <> "" Then
        lArrCnt = lArrCnt + 1
        ReDim Preserve ReturnArray(1 To lArrCnt) As String
        ReturnArray(lArrCnt) = OriginalString
    End If
    
    SplitStr = lArrCnt
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''   BELOW ARE THE EVENTS CALLED BY THE DATA PARSER   '''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Event_Unknown(strText)
'misc text recieved by battle.net, the text is not an event
If Left(strText, 15) = "Login incorrect" Then
AddChat "Your Login is Incorrect" & vbCrLf, vbRed
frmLogin.Show
        End If
End Sub

Private Sub Event_User(strUser, strFlag, strProduct)

    'Listing of a user in the current channel (after a Channel event)
    'strUser   : User's Name
    'strFlag   : User's Status Flag
    'strProduct: User's Product
    frmBot.ChannelUsers.ListItems.Add frmBot.ChannelUsers.ListItems.Count + 1, , strUser, , GetIconCode(strProduct, strFlag)
    If Mid(strFlag, 3, 1) = "1" Then frmBot.ChannelUsers.ListItems.Item(frmBot.ChannelUsers.ListItems.Count).ForeColor = vbYellow
End Sub

Private Sub Event_Join(strUser, strFlag, strProduct)
frmBot.ChannelUsers.ListItems.Add frmBot.ChannelUsers.ListItems.Count + 1, , strUser, , GetIconCode(strProduct, strFlag)
    If Mid(strFlag, 3, 1) = "1" Then frmBot.ChannelUsers.ListItems.Item(frmBot.ChannelUsers.ListItems.Count).ForeColor = vbYellow
End Sub


Private Sub Event_Leave(strUser, strFlag)
    'User leaves current channel
    'strUser: User's Name
    'strFlag: User's Status Flag
frmBot.ChannelUsers.ListItems.Remove (frmBot.ChannelUsers.FindItem(strUser).Index)
End Sub

Private Sub Event_RecvWhisper(strUser, strFlag, strText)
    AddChat "<From: " & strUser & "> ", vbYellow, strText & vbCrLf, &H808080
End Sub

Private Sub Event_Talk(strUser, strFlag, strText)
    'User Talks
    'strUser: User's Name
    'strFlag: User's Status Flag
    'strText: Whisper Text
   strText = ProfanityUnfilter(strText)
   If Mid(strFlag, 4, 2) = "2" Then
   AddChat "<" & strUser & "> ", vbWhite, strText & vbCrLf, vbWhite
   ElseIf Mid(strFlag, 4, 1) = "1" Then
AddChat "<" & strUser & "> ", 16776960, strText & vbCrLf, 16776960
   Else: AddChat "<" & strUser & "> ", vbYellow, strText & vbCrLf, vbWhite
   End If
ChatIn strUser, strText


End Sub

Private Sub Event_Broadcast(strText)
    'Broadcast from server administrator
    'strText: Broadcast Message
    AddChat strText & vbCrLf, vbRed
End Sub

Private Sub Event_Channel(strText)
    'Joined New Channel
    'strText: Channel Name
    frmBot.ChannelUsers.ListItems.Clear
    frmBot.ChannelUsers.ColumnHeaders(1).text = strText
    AddChat "Joined Channel: " & strText & vbCrLf, &HFF00&
    TheChannel = strText
    frmBot.Caption = "Vapor Bot - " & frmBot.strName & " in " & strText
End Sub

Private Sub Event_Flags(strUser, strFlag)

Dim newicon As Integer
newicon = ChangeFlags(strFlag)

    Dim i As Integer
    On Error Resume Next


For i = 0 To frmBot.ChannelUsers.ListItems.Count
        If frmBot.ChannelUsers.ListItems(i) = "" Then
            GoTo hi
        End If
        If frmBot.ChannelUsers.ListItems(i).text = strUser Then
                If strFlag = "0000" Then
        If frmBot.ChannelUsers.ListItems(i).SmallIcon = ICON_IGNORE Then GoTo muaha
        Exit Sub
        End If
        If strFlag = "0010" Then
        If frmBot.ChannelUsers.ListItems(i).SmallIcon = ICON_IGNORE Then GoTo muaha
        Exit Sub
        End If
muaha:
                frmBot.ChannelUsers.ListItems.Remove (i)
            Exit For
        End If
hi:
    Next i
frmBot.ChannelUsers.ListItems.Add (i), , strUser, , newicon
End Sub

Private Sub Event_SendWhisper(strUser, strFlag, strText)
AddChat "<To: " & strUser & "> ", &HFFFF00, strText & vbCrLf, &H808080
End Sub

Private Sub Event_ChannelFull(strText)
    'Channel is Full
    'strText: Name of Channel
    AddChat "Channel " & strText & " is full" & vbCrLf, &HFF00&
End Sub

Private Sub Event_ChannelNotExist(strText)
AddChat "Channel " & strText & " does not exist" & vbCrLf, &HFF00&
End Sub

Private Sub Event_ChannelRestricted(strText)
    'Channel is Restricted
    'strText: Name of Channel
    AddChat "Channel " & strText & " is restricted" & vbCrLf, &HFF00&
End Sub

Private Sub Event_Info(strText)
    'Information from the Server
    'strText: Info Message
    AddChat strText & vbCrLf, vbYellow
End Sub

Private Sub Event_Error(strText)
    'Error from the Server
    'strText: Error Message
    
    AddChat strText & vbCrLf, vbRed
End Sub

Private Sub Event_Emote(strUser, strFlag, strText)
    'Emote Event
    'strUser: User's Name
    'strFlag: User's Status Flag
    
    AddChat "<" & strUser & " " & strText & ">" & vbCrLf, vbYellow
End Sub

Private Sub Event_Name(strUser)
LoadDatabase
    frmBot.Caption = "Vapor Bot - " & strUser
    frmBot.strName = strUser
    AddChat "Successfully logged in as " & strUser & vbCrLf, vbGreen
End Sub
