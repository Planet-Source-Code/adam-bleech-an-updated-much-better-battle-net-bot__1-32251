Attribute VB_Name = "modCommands"
Public Sub LoadDatabase()
     'load text file into the array
     Dim MyString As String, strString As String
    On Error Resume Next
    Open App.Path & "\Users.txt" For Input As #1


    While Not EOF(1)
        Input #1, MyString$


        DoEvents
            strString = strString & MyString$ & vbCrLf
        Wend
        Close #1
        Close #1
        Close #1
        Database = Split(strString, vbCrLf)
    End Sub
Public Sub SaveDatabase()
    'save array as text file
    Dim SaveList As Long
    On Error Resume Next
    Open App.Path & "\Users.txt" For Output As #1


    For SaveList& = 0 To UBound(Database)
        Print #1, Database(SaveList&)
    Next SaveList&
    Close #1
End Sub
Public Function CheckDatabase(User) As Boolean
    'if the user is found in the database then the function
    'will return as true
    Dim i As Integer
    CheckDatabase = False
For i = 0 To UBound(Database)
If LCase(User) = LCase(Database(i)) Then CheckDatabase = True
    Next i
End Function

Public Sub ChatIn(Who, What)
'this little sub parses out all the commands and responds to
'them. :) You of course can add your own!
If CheckDatabase(Who) = False Then Exit Sub
Dim strChannel As String
If LCase(TheChannel) <> "the void" Then strChannel = TheChannel
If LCase(Left(What, 5)) = "!say " Then SendChat Mid(What, 6, Len(What) - 5)
If LCase(What) = "!server" Then SendChat "Connected to: " & IniServer & " " & frmBot.Winsock.RemoteHostIP
If LCase(What) = "!quit" Then End
If LCase(What) = "!rejoin" Then
SendChat "/join the void"
DoEvents
Pause (2)
SendChat "/join " & strChannel
End If
If LCase(What) = "!reconnect" Then
Battlenet.Disconnect frmBot.Winsock
Battlenet.Connect frmBot.Winsock
Pause (2)
SendChat "/join " & strChannel
End If
If LCase(Left(What, 5)) = "!add " Then
Pause (2)
ReDim Preserve Database(UBound(Database) + 1)
Database(UBound(Database)) = Mid(What, 6, Len(What) - 5)
SendChat "User " & Mid(What, 6, Len(What) - 5) & "has been added to database."
Pause (1)
SaveDatabase
End If
If LCase(Left(What, 8)) = "!remove " Then
    Dim i As Integer
For i = 0 To UBound(Database)
If LCase(Right(What, Len(What) - 8)) = LCase(Database(i)) Then
ArrayRemove Database, i
SendChat "User " & Right(What, Len(What) - 8) & " has been deleted."
SaveDatabase
    Exit Sub
    End If
    Next i
    SendChat "Error: No such user, " & Right(What, Len(What) - 8)
    End If
If LCase(Left(What, 6)) = "!join " Then SendChat "/join " & Right(What, Len(What) - 6)

End Sub
Private Sub ArrayCrunch(List)
'by James Vincent Carnicelli
    If UBound(List) = 0 Then
        On Error Resume Next
        List = Array()
    Else
        ReDim Preserve List(UBound(List) - 1)
    End If
End Sub
Private Sub ArraySetItem(List, ByVal Index As Integer, Item)
'by James Vincent Carnicelli
    If IsObject(Item) Then
        Set List(Index) = Item
    Else
        List(Index) = Item
    End If
End Sub
Public Sub ArrayRemove(List, ByVal Index As Integer)
'by James Vincent Carnicelli
Dim i As Integer
    For i = Index + 1 To UBound(List)
        ArraySetItem List, i - 1, List(i)
    Next
    ArrayCrunch List
End Sub
