VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Vapor Bot"
   ClientHeight    =   1245
   ClientLeft      =   3420
   ClientTop       =   5010
   ClientWidth     =   4005
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1245
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtpassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   0
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.TextBox txtlogin 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Battlenet.Login frmBot.Winsock, txtlogin.text, txtpassword.text
    Unload Me

End Sub

