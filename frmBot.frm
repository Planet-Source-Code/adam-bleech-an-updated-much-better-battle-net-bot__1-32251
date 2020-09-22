VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmBot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vapor Bot"
   ClientHeight    =   6060
   ClientLeft      =   1170
   ClientTop       =   4620
   ClientWidth     =   10785
   Icon            =   "frmBot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   10785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer AntiIdle 
      Interval        =   1000
      Left            =   2040
      Top             =   3720
   End
   Begin VB.CommandButton cmdWhisper 
      Caption         =   "Whisper"
      Height          =   255
      Left            =   9720
      TabIndex        =   4
      Top             =   5785
      Width           =   855
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   1560
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ChatIcons 
      Left            =   960
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBot.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBot.frx":0876
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBot.frx":0D62
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBot.frx":124E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBot.frx":173A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBot.frx":1C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBot.frx":2112
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBot.frx":25FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBot.frx":2AEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBot.frx":2FD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBot.frx":34C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBot.frx":39AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBot.frx":3DD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBot.frx":43CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBot.frx":4792
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBot.frx":5E0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBot.frx":743A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBot.frx":89E6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ChannelUsers 
      Height          =   5775
      Left            =   8520
      TabIndex        =   2
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   10186
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      SmallIcons      =   "ChatIcons"
      ForeColor       =   16777215
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   255
      Left            =   8760
      TabIndex        =   1
      Top             =   5785
      Width           =   855
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   5760
      Width           =   8535
   End
   Begin RichTextLib.RichTextBox txtChannel 
      Height          =   5775
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   10186
      _Version        =   393217
      BackColor       =   0
      TextRTF         =   $"frmBot.frx":A0AA
   End
   Begin VB.Menu mnuBot 
      Caption         =   "&Bot"
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "&Files"
      Begin VB.Menu mnuDatabase 
         Caption         =   "&UserDatabase"
      End
      Begin VB.Menu mnuConfig 
         Caption         =   "&Config"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh Files"
      End
   End
End
Attribute VB_Name = "frmBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strList As String   'Last Clicked Battle.net Name
Public strName As String   'Your assigned Battle.net name

Private Sub AntiIdle_Timer()
'notice this timers interval is 1000, that is for 1 second
IdleTick = IdleTick + 1 'this is a public variable, it adds a
'second to the variable each time the timer executes
If IdleTick >= 120 Then 'if 2 minutes have gone by then...
SendChat "/me is a VaporBot [Core]" 'send the idle message
IdleTick = 0 'reset the variable to start over
End If
End Sub

Private Sub ChannelUsers_Click()
strList = ChannelUsers.SelectedItem.text 'this stores you
'last clicked user to a variable to be used with whispers
End Sub

Private Sub cmdWhisper_Click()
SendChat "/w " & strList & " " & txtSend.text 'look above
End Sub

Private Sub Command1_Click()

End Sub



Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuConfig_Click()
frmEdit.Show
frmEdit.Caption = "Config.ini"
OpenFile "Config.ini"
End Sub

Private Sub mnuConnect_Click()
    'Connect to Battle.net
    If mnuConnect.Caption = "&Connect" Then
        Battlenet.Connect Winsock
        mnuConnect.Caption = "&Disconnect"
    Else
        Battlenet.Disconnect Winsock
        mnuConnect.Caption = "&Connect"
    End If
End Sub




Private Sub mnuDatabase_Click()
frmEdit.Show
frmEdit.Caption = "Users.txt"
OpenFile "Users.txt"
End Sub

Private Sub mnuRefresh_Click()
LoadDatabase
AddChat "UserDatabase Refreshed.", vbGreen
End Sub

Private Sub Winsock_Connect()
AddChat "Connected To Battle.net at " & Time & vbCrLf, vbGreen
Battlenet.Login frmBot.Winsock, IniUser, IniPass
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    'Battle.net Data Arrival
    Dim strTmp As String
    Winsock.GetData strTmp, vbString
    Battlenet.ParseData strTmp
End Sub

Private Sub Winsock_Close()
    AddChat "Server closed connection" & vbCrLf, vbRed
    Battlenet.Disconnect Winsock
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Winsock error
    AddChat Description & vbCrLf, vbRed
    Battlenet.Disconnect Winsock
End Sub

Private Sub mnuExit_Click()
    'Disconnect and Exit
    Battlenet.Disconnect Winsock
    End
End Sub
Private Sub cmdSend_Click()
Battlenet2.SendChat txtSend.text
DoEvents
txtSend.text = ""
End Sub


