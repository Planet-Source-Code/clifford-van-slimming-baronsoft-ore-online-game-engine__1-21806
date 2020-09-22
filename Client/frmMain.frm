VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ORE Client"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11895
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   552
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   793
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   9390
      Top             =   450
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   2048
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
   Begin VB.TextBox DrpAmountTxt 
      Height          =   285
      Left            =   8400
      TabIndex        =   17
      Text            =   "1"
      Top             =   5580
      Width           =   855
   End
   Begin VB.CommandButton DropCmd 
      Caption         =   "Drop"
      Height          =   315
      Left            =   8400
      TabIndex        =   15
      Top             =   5220
      Width           =   855
   End
   Begin VB.CommandButton GetCmd 
      Caption         =   "Get"
      Height          =   315
      Left            =   9300
      TabIndex        =   14
      Top             =   5220
      Width           =   855
   End
   Begin VB.CommandButton UseCmd 
      Caption         =   "Use"
      Height          =   315
      Left            =   10200
      TabIndex        =   13
      Top             =   5220
      Width           =   855
   End
   Begin VB.ListBox ObjLst 
      Height          =   2985
      Left            =   8400
      TabIndex        =   12
      Top             =   2220
      Width           =   2655
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1755
      Left            =   8400
      ScaleHeight     =   117
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   2
      Top             =   6360
      Width           =   3255
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mana:"
         Height          =   225
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   465
      End
      Begin VB.Shape HPShp 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   150
         Left            =   720
         Top             =   540
         Width           =   2250
      End
      Begin VB.Shape STAShp 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   150
         Left            =   720
         Top             =   750
         Width           =   2250
      End
      Begin VB.Shape MANShp 
         BackColor       =   &H00FFFF00&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   150
         Left            =   720
         Top             =   960
         Width           =   2250
      End
      Begin VB.Label GldLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   750
         TabIndex        =   9
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label LvlLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   2190
         TabIndex        =   8
         Top             =   240
         Width           =   75
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "STA:"
         Height          =   225
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "MP:"
         Height          =   225
         Left            =   -390
         TabIndex        =   5
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Gold:"
         Height          =   225
         Left            =   150
         TabIndex        =   4
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Level:"
         Height          =   225
         Left            =   1590
         TabIndex        =   3
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Timer FPSTimer 
      Interval        =   1000
      Left            =   8400
      Top             =   480
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   8175
   End
   Begin MCI.MMControl MidiPlayer 
      Height          =   420
      Left            =   8880
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "Sequencer"
      FileName        =   ""
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1485
      Left            =   60
      TabIndex        =   11
      Top             =   390
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   2619
      _Version        =   393217
      BackColor       =   16777215
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label MainViewLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Downloading New Map: 0%"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   840
      TabIndex        =   18
      Top             =   4860
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Shape MainViewShp 
      Height          =   6240
      Left            =   60
      Top             =   1980
      Width           =   8160
   End
   Begin VB.Label MapNameLbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   345
      Left            =   8400
      TabIndex        =   0
      Top             =   60
      Width           =   3285
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit








Private Sub DropCmd_Click()

'Send the drop command
If ObjLst.ListIndex > -1 Then
    SendData "DRP" & ObjLst.ListIndex + 1 & "," & DrpAmountTxt.Text
End If

End Sub

Private Sub DrpAmountTxt_Change()

'Make sure amount is legal
If DrpAmountTxt.Text < 1 Then
    DrpAmountTxt.Text = MAX_INVENTORY_OBJS
End If

If DrpAmountTxt.Text > MAX_INVENTORY_OBJS Then
    DrpAmountTxt.Text = 1
End If

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

'If a letter,number,or backspace key send it to the sendtxt box
If KeyAscii >= 32 And KeyAscii <= 126 Or KeyAscii = 8 Then
    SendTxt.SetFocus
    Exit Sub
End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'Attack key
If KeyCode = 18 Then
    SendData ("ATT")
    KeyCode = 0
    Exit Sub
End If

End Sub


Private Sub Form_Load()

'Update main caption
frmMain.Caption = frmMain.Caption & " V " & App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*****************************************************************
'See if user is clicking in the view window then send
'the tile click position to the server
'*****************************************************************
Dim tX As Integer
Dim tY As Integer

'Make sure engine is running
If EngineRun = False Then Exit Sub

'Don't do if downloading map
If DownloadingMap = True Then
    Exit Sub
End If

'Make sure click is in view window
If X <= MainViewShp.Left Or X >= MainViewShp.Left + MainViewWidth Or Y <= MainViewShp.Top Or Y >= MainViewShp.Top + MainViewHeight Then
    Exit Sub
End If

ConvertCPtoTP MainViewShp.Left, MainViewShp.Top, X, Y, tX, tY

If Button = vbLeftButton Then
    SendData "LC" & tX & "," & tY
Else
    SendData "RC" & tX & "," & tY
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'Allow the MainLoop to close program
If prgRun = True Then
    prgRun = False
    Cancel = 1
End If

End Sub

Private Sub FPSTimer_Timer()

'Display and reset FPS
FramesPerSec = FramesPerSecCounter
FramesPerSecCounter = 0

End Sub









Private Sub MainPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub




Private Sub GetCmd_Click()

'Send the get command
SendData "GET"

End Sub

Private Sub MidiPlayer_StatusUpdate()

'LSee if MIDI is done
If MidiPlayer.Length = MidiPlayer.Position Then
        
    'Loop if needed
    If LoopMidi Then
        Call PlayMidi(CurMidi)
    End If

End If

End Sub


Private Sub ObjLst_DblClick()

Call UseCmd_Click

End Sub

Private Sub ObjLst_KeyDown(KeyCode As Integer, Shift As Integer)

    SendTxt.SetFocus

End Sub


Private Sub RecTxt_Change()

    SendTxt.SetFocus

End Sub

Private Sub SendTxt_Change()

stxtbuffer = SendTxt.Text

End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)

'BackSpace
If KeyAscii = 8 Then
    Exit Sub
End If

'Every other letter
If KeyAscii >= 32 And KeyAscii <= 126 Then
    Exit Sub
End If

KeyAscii = 0

End Sub


Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)

Dim retcode As Integer

'Send text
If KeyCode = vbKeyReturn Then

    'Command
    If UCase(stxtbuffer) = "/MIDIOFF" Then
        retcode = mciSendString("close all", 0, 0, 0)
        LoopMidi = 0
        
    ElseIf Left$(stxtbuffer, 1) = "/" Then
        SendData (stxtbuffer)

    'yell
    ElseIf Left$(stxtbuffer, 1) = "'" Then
        SendData ("'" & Right$(stxtbuffer, Len(stxtbuffer) - 1))
    
    'Shout
    ElseIf Left$(stxtbuffer, 1) = "-" Then
        SendData ("-" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

    'Whisper
    ElseIf Left$(stxtbuffer, 1) = "\" Then
        SendData ("\" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

    'Emote
    ElseIf Left$(stxtbuffer, 1) = ":" Then
        SendData (":" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

    'Say
    ElseIf stxtbuffer <> "" Then
        SendData (";" & stxtbuffer)

    End If

    stxtbuffer = ""
    SendTxt.Text = ""
    KeyCode = 0
    Exit Sub

End If

End Sub

Private Sub Socket1_Connect()

Call Login
Call SetConnected

End Sub


Private Sub Socket1_Disconnect()

prgRun = False
Connected = False

End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
'*********************************************
'Handle socket errors
'*********************************************

Select Case (ErrorCode)

Case 24065
    MsgBox "The server seems to be down or unreachable. Please try again."
    frmConnect.MousePointer = 1
    Response = 0
    
Case 24061
    MsgBox "The server seems to be down or unreachable. Please try again."
    frmConnect.MousePointer = 1
    Response = 0
    
Case 24064
    MsgBox "The server seems to be down or unreachable. Please try again."
    frmConnect.MousePointer = 1
    Response = 0
    
Case Else
    MsgBox (ErrorString)
    frmConnect.MousePointer = 1

End Select


End Sub

Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
'*********************************************
'Seperate lines by ENDC and send each to HandleData()
'*********************************************

Dim LoopC As Integer

Dim RD As String
Dim rBuffer(1 To 500) As String
Static TempString As String

Dim CR As Integer
Dim tChar As String
Dim sChar As Integer
Dim eChar As Integer

Socket1.Read RD, DataLength

'Check for previous broken data and add to current data
If TempString <> "" Then
    RD = TempString & RD
    TempString = ""
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
    TempString = Mid$(RD, sChar, Len(RD))
End If

'Send buffer to Handle data
For LoopC = 1 To CR
    Call HandleData(rBuffer(LoopC))
Next LoopC

End Sub


Private Sub UseCmd_Click()

'Send use command
If ObjLst.ListIndex > -1 Then
    SendData "USE" & ObjLst.ListIndex + 1
End If

End Sub


