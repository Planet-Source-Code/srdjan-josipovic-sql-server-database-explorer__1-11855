VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "SQL Database Explorer"
   ClientHeight    =   4800
   ClientLeft      =   3390
   ClientTop       =   3450
   ClientWidth     =   6435
   LinkTopic       =   "Form2"
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   4800
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtDatabase 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USER ID:"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   480
      TabIndex        =   8
      Top             =   120
      Width           =   930
   End
   Begin VB.Label lblPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD:"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1260
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SERVER:"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   480
      TabIndex        =   6
      Top             =   1080
      Width           =   930
   End
   Begin VB.Label lblDatabase 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATABASE:"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1170
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConnect_Click()
On Error GoTo EH
   frmMain.Show
   Me.Hide
Exit Sub

EH:
   Unload frmMain
   Me.Show
End Sub

Private Sub cmdExit_Click()
   End
End Sub

