VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ORE Server"
   ClientHeight    =   4800
   ClientLeft      =   1950
   ClientTop       =   1530
   ClientWidth     =   3645
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4800
   ScaleWidth      =   3645
   WindowState     =   1  'Minimized
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   1830
      Top             =   3540
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin SocketWrenchCtrl.Socket Socket2 
      Index           =   0
      Left            =   2280
      Top             =   3540
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer AutoMaptimer 
      Interval        =   5
      Left            =   2760
      Top             =   3540
   End
   Begin VB.ListBox Userslst 
      Height          =   2160
      Left            =   120
      TabIndex        =   6
      Top             =   1260
      Width           =   3435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reset Server"
      Height          =   510
      Left            =   120
      TabIndex        =   5
      Top             =   3540
      Width           =   1635
   End
   Begin VB.TextBox LocalAdd 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.Timer GameTimer 
      Interval        =   50
      Left            =   3180
      Top             =   3540
   End
   Begin VB.TextBox txPortNumber 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txStatus 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   4380
      Width           =   3375
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1560
      TabIndex        =   8
      Top             =   4140
      Width           =   450
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Users Online"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   7
      Top             =   1020
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Server IP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   300
      Width           =   705
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Running on Port:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   15
      TabIndex        =   1
      Top             =   660
      Width           =   1200
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub AutoMaptimer_Timer()
'*****************************************************************
'Send out map updates if needed
'*****************************************************************
Dim UserIndex As Integer
Dim LoopC As Integer

For UserIndex = 1 To LastUser
    If UserList(UserIndex).Flags.UserLogged = 1 Then
    
        'Send map chunk
        If UserList(UserIndex).Flags.DownloadingMap = 1 Then
            For LoopC = 1 To 15
                If UserList(UserIndex).Flags.DownloadingMap Then
                    SendNextMapTile UserIndex
                End If
            Next LoopC
        End If
        
    End If
Next UserIndex

End Sub

Private Sub Command1_Click()

Call Restart

End Sub


Private Sub Form_Load()
'*****************************************************************
'Load up server
'*****************************************************************
Dim LoopC As Integer

'*** Init vars ***
frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
ENDL = Chr(13) & Chr(10)
ENDC = Chr(1)
IniPath = App.Path & "\"
CharPath = App.Path & "\Charfile\"

'Setup Map borders
MinXBorder = XMinMapSize + (XWindow \ 2)
MaxXBorder = XMaxMapSize - (XWindow \ 2)
MinYBorder = YMinMapSize + (YWindow \ 2)
MaxYBorder = YMaxMapSize - (YWindow \ 2)

'Reset User connections
For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
Next LoopC


'*** Load data ***

Call LoadSini
Call LoadMapData
Call LoadOBJData

'*** Setup sockets ***

frmMain.Socket1.AddressFamily = AF_INET
frmMain.Socket1.Protocol = IPPROTO_IP
frmMain.Socket1.SocketType = SOCK_STREAM
frmMain.Socket1.Binary = False
frmMain.Socket1.Blocking = False
frmMain.Socket1.BufferSize = SOCKET_BUFFER_SIZE

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Binary = False
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = SOCKET_BUFFER_SIZE

'*** Listen ***
frmMain.Socket1.LocalPort = Val(frmMain.txPortNumber.Text)
frmMain.Socket1.Listen
  
'*** Misc ***
'Hide
If HideMe = 1 Then
    frmMain.Hide
End If

'Show status
frmMain.txStatus.Text = "Listening for connection ..."
Call RefreshUserListBox

'Show local IP
frmMain.LocalAdd.Text = frmMain.Socket1.LocalAddress

'Log it
Open App.Path & "\Main.log" For Append Shared As #5
Print #5, "Server started. " & Time & " " & Date
Close #5

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim LoopC As Integer

'ensure that the sockets are closed, ignore any errors
On Error Resume Next

Socket1.Cleanup

For LoopC = 1 To MaxUsers
    CloseSocket (LoopC)
Next

'Log it
Open App.Path & "\Main.log" For Append Shared As #5
Print #5, "Server unloaded. " & Time & " " & Date
Close #5

End

End Sub













Sub GameTimer_Timer()
'*****************************************************************
'update world
'*****************************************************************
Dim UserIndex As Integer
Dim NPCIndex As Integer
Dim TempPos As WorldPos

'Update Users
For UserIndex = 1 To LastUser

    'make sure user is logged on
    If UserList(UserIndex).Flags.UserLogged = 1 Then
        
        'Do special tile events
        Call DoTileEvents(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
        
        'Update HP
        UserList(UserIndex).Counters.HPCounter = UserList(UserIndex).Counters.HPCounter + 1
        If UserList(UserIndex).Counters.HPCounter >= STAT_METRATE Then
            AddtoVar UserList(UserIndex).Stats.MinHP, UserList(UserIndex).Stats.MET, UserList(UserIndex).Stats.MaxHP
            UserList(UserIndex).Counters.HPCounter = 0
            UserList(UserIndex).Flags.StatsChanged = 1
        End If

        'Update HP
        UserList(UserIndex).Counters.STACounter = UserList(UserIndex).Counters.STACounter + 1
        If UserList(UserIndex).Counters.STACounter >= STAT_FITRATE Then
            AddtoVar UserList(UserIndex).Stats.MinSTA, UserList(UserIndex).Stats.FIT, UserList(UserIndex).Stats.MaxSTA
            UserList(UserIndex).Counters.STACounter = 0
            UserList(UserIndex).Flags.StatsChanged = 1
        End If

        'Update attack counter
        If UserList(UserIndex).Counters.AttackCounter > 0 Then
            UserList(UserIndex).Counters.AttackCounter = UserList(UserIndex).Counters.AttackCounter - 1
        End If
        
        'Update Stats box if need be
        If UserList(UserIndex).Flags.StatsChanged Then
            SendUserStatsBox UserIndex
            UserList(UserIndex).Flags.StatsChanged = 0
        End If
            
        'Update idle counter
        UserList(UserIndex).Counters.IdleCount = UserList(UserIndex).Counters.IdleCount + 1
        If UserList(UserIndex).Counters.IdleCount >= IdleLimit Then
            Call SendData(ToIndex, UserIndex, 0, "!!Sorry you have been idle to long. Disconnected..")
            Call CloseSocket(UserIndex)
        End If
            
    End If

Next UserIndex

'Update NPCs
For NPCIndex = 1 To LastNPC

    'make sure NPC is active
    If NPCList(NPCIndex).Flags.NPCActive = 1 Then
        
        'Only update npcs in user populated maps
        If MapInfo(NPCList(NPCIndex).StartPos.Map).NumUsers > 0 Then

            'see if npc is alive
            If NPCList(NPCIndex).Flags.NPCAlive Then
                'randomly call NPCAI
                If Int(RandomNumber(1, 30)) = 1 Then
                    Call NPCAI(NPCIndex)
                End If
            Else
            
                'update respawncounter
                If NPCList(NPCIndex).RespawnWait > 0 Then
                    NPCList(NPCIndex).Counters.RespawnCounter = NPCList(NPCIndex).Counters.RespawnCounter - 1
                    If NPCList(NPCIndex).Counters.RespawnCounter <= 0 Then
                        SpawnNPC NPCIndex, NPCList(NPCIndex).StartPos.Map, NPCList(NPCIndex).StartPos.X, NPCList(NPCIndex).StartPos.Y
                    End If
                End If
                
            End If
            
        End If
    
    End If

Next NPCIndex

End Sub

Sub Socket1_Accept(SocketId As Integer)
'*********************************************
'Accepts new user and assigns an open Index
'*********************************************
Dim Index As Integer

Index = NextOpenUser

If UserList(Index).ConnID >= 0 Then
    'Close down user socket
    Call CloseSocket(Index)
End If

UserList(Index).ConnID = SocketId
Load Socket2(Index)

Socket2(Index).AddressFamily = AF_INET
Socket2(Index).Protocol = IPPROTO_IP
Socket2(Index).SocketType = SOCK_STREAM
Socket2(Index).Binary = False
Socket2(Index).BufferSize = SOCKET_BUFFER_SIZE
Socket2(Index).Blocking = False

Socket2(Index).Accept = SocketId

End Sub

Sub Socket2_Disconnect(Index As Integer)
'*********************************************
'Begins close procedure
'*********************************************

CloseSocket (Index)

End Sub


Sub Socket2_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)
'*********************************************
'Seperate lines by ENDC and send each to HandleData()
'*********************************************
Dim LoopC As Integer
Dim RD As String
Dim rBuffer(1 To COMMAND_BUFFER_SIZE) As String
Dim CR As Integer
Dim tChar As String
Dim sChar As Integer
Dim eChar As Integer

Socket2(Index).Read RD, DataLength

'Check for previous broken data and add to current data
If UserList(Index).RDBuffer <> "" Then
    RD = UserList(Index).RDBuffer & RD
    UserList(Index).RDBuffer = ""
End If

'Check for more than one line
sChar = 1
For LoopC = 1 To Len(RD)

    tChar = Mid$(RD, LoopC, 1)

    If tChar = ENDC Then
        CR = CR + 1
        eChar = LoopC - sChar
        rBuffer(CR) = Mid$(RD, sChar, eChar)
        sChar = LoopC + 1
    End If
        
Next LoopC

'Check for broken line and save for next time
If Len(RD) - (sChar - 1) <> 0 Then
    UserList(Index).RDBuffer = Mid$(RD, sChar, Len(RD))
End If

'Send buffer to Handle data
For LoopC = 1 To CR
    Call HandleData(Index, rBuffer(LoopC))
Next LoopC

End Sub


