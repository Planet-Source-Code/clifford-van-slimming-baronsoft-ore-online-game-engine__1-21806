VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Connect to Server"
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   3900
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox IPTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   1260
      TabIndex        =   11
      Text            =   "localhost"
      Top             =   2880
      Width           =   2595
   End
   Begin VB.TextBox PortTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   1260
      TabIndex        =   9
      Text            =   "7777"
      Top             =   2490
      Width           =   2595
   End
   Begin VB.CheckBox SavePassChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1260
      TabIndex        =   7
      Top             =   1770
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      TabIndex        =   6
      Top             =   3870
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1410
      TabIndex        =   5
      Top             =   3870
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2670
      TabIndex        =   4
      Top             =   3870
      Width           =   1095
   End
   Begin VB.TextBox PasswordTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1350
      Width           =   2595
   End
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   1260
      TabIndex        =   0
      Top             =   900
      Width           =   2595
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connect to Server"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   660
      TabIndex        =   13
      Top             =   240
      Width           =   2430
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name/IP"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   90
      TabIndex        =   12
      Top             =   2850
      Width           =   885
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   90
      TabIndex        =   10
      Top             =   2520
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1500
      TabIndex        =   8
      Top             =   1740
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   60
      TabIndex        =   3
      Top             =   1470
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Top             =   990
      Width           =   600
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'*****************************************************************
'Makes sure user data is ok then trys to connect to server
'*****************************************************************

On Error Resume Next

If frmConnect.MousePointer = 11 Then
    Exit Sub
End If

'update user info
UserName = NameTxt.Text
UserPassword = PasswordTxt.Text
UserServerIP = IPTxt.Text
UserPort = Val(PortTxt.Text)

If CheckUserData = True Then
       
    'FrmMain.Socket1.Close
    frmMain.Socket1.HostName = UserServerIP
    frmMain.Socket1.RemotePort = UserPort

    SendNewChar = False
    frmConnect.MousePointer = 11
    frmMain.Socket1.Connect
    
End If

End Sub

Private Sub Command2_Click()
'*****************************************************************
'Makes sure user data is ok then begins new character process
'*****************************************************************

'update user info
UserName = NameTxt.Text
UserPassword = PasswordTxt.Text
UserServerIP = IPTxt.Text
UserPort = Val(PortTxt.Text)

UserBody = 1
UserHead = 1

If CheckUserData = True Then

    'FrmMain.Socket1.Close
    frmMain.Socket1.HostName = UserServerIP
    frmMain.Socket1.RemotePort = UserPort

    SendNewChar = True
    frmConnect.MousePointer = 11
    frmMain.Socket1.Connect

End If

End Sub


Private Sub Command3_Click()

'update info
UserName = NameTxt.Text
UserPassword = PasswordTxt.Text
UserServerIP = IPTxt.Text
UserPort = Val(PortTxt.Text)

Call SaveGameini

frmConnect.MousePointer = 1
frmMain.MousePointer = 1

'End program
prgRun = False

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'Make Server IP and Port box visible
If KeyCode = vbKeyI And Shift = vbCtrlMask Then
    
    'Port
    PortTxt.Visible = True
    Label4.Visible = True
    
    'Server IP
    IPTxt.Text = "localhost"
    IPTxt.Visible = True
    Label5.Visible = True
    
    KeyCode = 0
    Exit Sub
End If

End Sub

Private Sub Form_Load()

'Get Game.ini Data
If FileExist(IniPath & "Game.ini", vbNormal) = True Then
    NameTxt.Text = GetVar(IniPath & "Game.ini", "INIT", "Name")
    PasswordTxt.Text = GetVar(IniPath & "Game.ini", "INIT", "Password")
    PortTxt.Text = GetVar(IniPath & "Game.ini", "INIT", "Port")
End If

End Sub


