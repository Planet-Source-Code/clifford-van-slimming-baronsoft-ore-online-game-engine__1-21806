VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ORE Map Editor"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11895
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   552
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   793
   Visible         =   0   'False
   Begin VB.TextBox MapNameTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   8340
      TabIndex        =   43
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox MapVersionTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   10920
      TabIndex        =   39
      Top             =   7080
      Width           =   795
   End
   Begin VB.TextBox MusNumTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   10920
      TabIndex        =   38
      Top             =   6600
      Width           =   795
   End
   Begin VB.TextBox StartPosTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   6120
      Width           =   795
   End
   Begin VB.TextBox OBJAmountTxt 
      Height          =   285
      Left            =   9240
      TabIndex        =   35
      Text            =   "1"
      Top             =   7320
      Width           =   555
   End
   Begin VB.CheckBox EraseObjChk 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Erase OBJ"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8340
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   7620
      Width           =   1215
   End
   Begin VB.CommandButton PlaceObjCmd 
      Caption         =   "Place OBJ"
      Height          =   255
      Left            =   8340
      TabIndex        =   33
      Top             =   7860
      Width           =   1515
   End
   Begin VB.ListBox ObjLst 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1230
      Left            =   8340
      TabIndex        =   32
      Top             =   6060
      Width           =   1575
   End
   Begin VB.CheckBox EraseNPCChk 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Erase NPC"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   10110
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton PlaceNPCCmd 
      Caption         =   "Place NPC"
      Height          =   255
      Left            =   10080
      TabIndex        =   30
      Top             =   5640
      Width           =   1515
   End
   Begin VB.ListBox NPCLst 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1620
      Left            =   10080
      TabIndex        =   29
      Top             =   3750
      Width           =   1575
   End
   Begin VB.TextBox YExitTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   9060
      TabIndex        =   25
      Text            =   "1"
      Top             =   4680
      Width           =   795
   End
   Begin VB.TextBox XExitTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   9060
      TabIndex        =   24
      Text            =   "1"
      Top             =   4200
      Width           =   795
   End
   Begin VB.TextBox MapExitTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   9060
      TabIndex        =   23
      Text            =   "1"
      Top             =   3780
      Width           =   795
   End
   Begin VB.CheckBox EraseExitChk 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Erase Exit"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   8340
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5340
      Width           =   1215
   End
   Begin VB.CommandButton PlaceExitCmd 
      Caption         =   "Place Exit"
      Height          =   255
      Left            =   8340
      TabIndex        =   21
      Top             =   5640
      Width           =   1515
   End
   Begin VB.CheckBox WalkModeChk 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Walk Mode"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   10050
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Show Blocked Tiles"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   10050
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1380
      Width           =   1815
   End
   Begin VB.CheckBox DrawGridChk 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Draw Grid"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   10050
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1155
   End
   Begin VB.PictureBox ShowPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   1800
      Left            =   4380
      ScaleHeight     =   118
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   254
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   120
      Width           =   3840
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "frmMain.frx":0442
      Top             =   330
      Width           =   2775
   End
   Begin VB.CommandButton PlaceGrhCmd 
      Caption         =   "Place Grh"
      Enabled         =   0   'False
      Height          =   255
      Left            =   8340
      TabIndex        =   13
      Top             =   3300
      Width           =   1515
   End
   Begin VB.CommandButton PlaceBlockCmd 
      Caption         =   "Change Blocked"
      Height          =   255
      Left            =   10080
      TabIndex        =   12
      Top             =   3300
      Width           =   1515
   End
   Begin VB.TextBox Grhtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   8340
      TabIndex        =   9
      Text            =   "1"
      Top             =   1380
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Up"
      Height          =   255
      Left            =   9180
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1260
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Down"
      Height          =   255
      Left            =   9180
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1620
      Width           =   615
   End
   Begin VB.TextBox Layertxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   8340
      TabIndex        =   6
      Text            =   "1"
      Top             =   2160
      Width           =   555
   End
   Begin VB.CheckBox Blockedchk 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Blocked"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   10200
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2940
      Width           =   915
   End
   Begin VB.CheckBox EraseAllchk 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Erase All"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8340
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2940
      Width           =   1335
   End
   Begin VB.CheckBox Erasechk 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Erase Layer"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8340
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Timer FPSTimer 
      Interval        =   1000
      Left            =   9720
      Top             =   600
   End
   Begin VB.ListBox MapLst 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1590
      Left            =   3060
      TabIndex        =   0
      Top             =   330
      Width           =   1095
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "MusNum:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   10020
      TabIndex        =   42
      Top             =   6690
      Width           =   840
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   10020
      TabIndex        =   41
      Top             =   7140
      Width           =   690
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "StartPos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   10020
      TabIndex        =   40
      Top             =   6210
      Width           =   810
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   8400
      TabIndex        =   36
      Top             =   7320
      Width           =   675
   End
   Begin VB.Shape Shape5 
      BorderWidth     =   2
      Height          =   2175
      Left            =   8280
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Shape MainViewShp 
      Height          =   6240
      Left            =   60
      Top             =   1950
      Width           =   8160
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   2
      Height          =   2295
      Left            =   10020
      Top             =   3660
      Width           =   1695
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   2295
      Left            =   8280
      Top             =   3660
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   2535
      Left            =   8280
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   795
      Left            =   10020
      Top             =   2820
      Width           =   1635
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "MAP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   8340
      TabIndex        =   28
      Top             =   3840
      Width           =   570
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   8370
      TabIndex        =   27
      Top             =   4710
      Width           =   225
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   8370
      TabIndex        =   26
      Top             =   4290
      Width           =   225
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maps:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3300
      TabIndex        =   16
      Top             =   60
      Width           =   645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Info:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1020
      TabIndex        =   15
      Top             =   60
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Layer"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8340
      TabIndex        =   11
      Top             =   1950
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Grh"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8340
      TabIndex        =   10
      Top             =   1140
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FPS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   10260
      TabIndex        =   2
      Top             =   660
      Width           =   600
   End
   Begin VB.Label FPSLbl 
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
      Height          =   420
      Left            =   10920
      TabIndex        =   1
      Top             =   600
      Width           =   795
   End
   Begin VB.Menu FileMnu 
      Caption         =   "File"
      Begin VB.Menu SaveMnu 
         Caption         =   "Save"
      End
      Begin VB.Menu SaveNewMnu 
         Caption         =   "Save as New Map"
      End
   End
   Begin VB.Menu OptionMnu 
      Caption         =   "Options"
      Begin VB.Menu ClsRoomMnu 
         Caption         =   "Clear Map"
      End
      Begin VB.Menu ClsBordMnu 
         Caption         =   "Clear Border"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Blockedchk_Click()

Call PlaceBlockCmd_Click

End Sub

Private Sub Check1_Click()

If DrawBlock = True Then
    DrawBlock = False
Else
    DrawBlock = True
End If

End Sub

Private Sub ClsBordMnu_Click()

'*****************************************************************
'Clears a border in a room with current GRH
'*****************************************************************

Dim Y As Integer
Dim X As Integer

If CurMap = 0 Then
    Exit Sub
End If

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then

            MapData(X, Y).Graphic(Val(Layertxt.Text)).GrhIndex = Val(frmMain.Grhtxt)

            'Setup GRH for layer

            InitGrh MapData(X, Y).Graphic(Val(Layertxt.Text)), Val(Grhtxt.Text)

            'Erase NPCs
            If MapData(X, Y).NPCIndex > 0 Then
                EraseChar MapData(X, Y).CharIndex
                MapData(X, Y).NPCIndex = 0
            End If

            'Erase Objs
            MapData(X, Y).OBJInfo.OBJIndex = 0
            MapData(X, Y).OBJInfo.Amount = 0
            MapData(X, Y).ObjGrh.GrhIndex = 0

            'Clear exits
            MapData(X, Y).TileExit.Map = 0
            MapData(X, Y).TileExit.X = 0
            MapData(X, Y).TileExit.Y = 0

        End If

    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1

End Sub

Private Sub ClsRoomMnu_Click()
'*****************************************************************
'Clears all layers
'*****************************************************************

Dim Y As Integer
Dim X As Integer

If CurMap = 0 Then
    Exit Sub
End If

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        'Change blockes status
        MapData(X, Y).Blocked = Blockedchk.value

        'Erase layer 2 and 4
        MapData(X, Y).Graphic(2).GrhIndex = 0
        MapData(X, Y).Graphic(3).GrhIndex = 0
        MapData(X, Y).Graphic(4).GrhIndex = 0

        'Erase NPCs
        If MapData(X, Y).NPCIndex > 0 Then
            EraseChar MapData(X, Y).CharIndex
            MapData(X, Y).NPCIndex = 0
        End If

        'Erase Objs
        MapData(X, Y).OBJInfo.OBJIndex = 0
        MapData(X, Y).OBJInfo.Amount = 0
        MapData(X, Y).ObjGrh.GrhIndex = 0

        'Clear exits
        MapData(X, Y).TileExit.Map = 0
        MapData(X, Y).TileExit.X = 0
        MapData(X, Y).TileExit.Y = 0

        'Place layer 1
        MapData(X, Y).Graphic(1).GrhIndex = Val(frmMain.Grhtxt)

        'Setup GRH for layer 1
        InitGrh MapData(X, Y).Graphic(1), Val(Grhtxt.Text)

    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1

End Sub


Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim Grh As Integer

'Set Place GRh mode
Call PlaceGrhCmd_Click

Grh = Val(Grhtxt.Text)

'Add to current Grh number
If Button = vbLeftButton Then
   Grh = Grh + 1
End If

If Button = vbRightButton Then
    Grh = Grh + 10
End If

'Update Grhtxt
Grhtxt.Text = Grh
Grh = Val(Grhtxt)

'If blank find next valid Grh
If GrhData(Grh).NumFrames = 0 Then
    
    Do Until GrhData(Grh).NumFrames > 0
        Grh = Grh + 1
        If Grh > NumGrhs Then
            Grh = 1
        End If
    Loop
    
End If

'Update Grhtxt
Grhtxt.Text = Grh

End Sub


Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim Grh As Integer

'Set Place GRh mode
Call PlaceGrhCmd_Click

Grh = Val(Grhtxt.Text)

'Add to current Grh number
If Button = vbLeftButton Then
   Grh = Grh - 1
End If

If Button = vbRightButton Then
    Grh = Grh - 10
End If

'Update Grhtxt
Grhtxt.Text = Grh
Grh = Val(Grhtxt)

'If blank find next valid Grh
If GrhData(Grh).NumFrames = 0 Then
    
    Do Until GrhData(Grh).NumFrames > 0
        Grh = Grh - 1
        If Grh < 1 Then
            Grh = NumGrhs
        End If
    Loop
    
End If

'Update Grhtxt
Grhtxt.Text = Grh

End Sub








Private Sub DrawGridChk_Click()

If DrawGrid = True Then
    DrawGrid = False
Else
    DrawGrid = True
End If

End Sub

Private Sub EraseAllchk_Click()

'Set Place GRh mode
Call PlaceGrhCmd_Click

Erasechk.value = False


End Sub

Private Sub Erasechk_Click()

'Set Place GRh mode
Call PlaceGrhCmd_Click

EraseAllchk.value = False

End Sub

Private Sub EraseExitChk_Click()

Call PlaceExitCmd_Click

End Sub

Private Sub EraseNPCChk_Click()

Call PlaceNPCCmd_Click

End Sub

Private Sub EraseObjChk_Click()

Call PlaceObjCmd_Click

End Sub

Private Sub Form_Load()

'Update main caption
frmMain.Caption = frmMain.Caption & " V " & App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim tX As Integer
Dim tY As Integer

'Make sure map is loaded
If CurMap <= 0 Then Exit Sub

'Make sure click is in view window
If X <= MainViewShp.Left Or X >= MainViewShp.Left + MainViewWidth Or Y <= MainViewShp.Top Or Y >= MainViewShp.Top + MainViewHeight Then
    Exit Sub
End If

ConvertCPtoTP MainViewShp.Left, MainViewShp.Top, X, Y, tX, tY

ReacttoMouseClick Button, tX, tY

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim tX As Integer
Dim tY As Integer

'Make sure map is loaded
If CurMap <= 0 Then Exit Sub

'Make sure click is in view window
If X <= MainViewShp.Left Or X >= MainViewShp.Left + MainViewWidth Or Y <= MainViewShp.Top Or Y >= MainViewShp.Top + MainViewHeight Then
    Exit Sub
End If

ConvertCPtoTP MainViewShp.Left, MainViewShp.Top, X, Y, tX, tY

ReacttoMouseClick Button, tX, tY

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'Allow MainLoop to close program
If prgRun = True Then
    prgRun = False
    Cancel = 1
End If

End Sub

Private Sub FPSTimer_Timer()

'Display and reset FPS
FramesPerSec = FramesPerSecCounter
FramesPerSecCounter = 0

FPSLbl.Caption = FramesPerSec

End Sub




Private Sub Grhtxt_Change()

If Val(Grhtxt.Text) < 1 Then
  Grhtxt.Text = NumGrhs
  Exit Sub
End If

If Val(Grhtxt.Text) > NumGrhs Then
  Grhtxt.Text = 1
  Exit Sub
End If

'Change CurrentGrh
CurrentGrh.GrhIndex = Val(frmMain.Grhtxt.Text)
CurrentGrh.Started = 1
CurrentGrh.FrameCounter = 1
CurrentGrh.SpeedCounter = GrhData(CurrentGrh.GrhIndex).Speed

End Sub


Private Sub Layertxt_Change()

If Val(Layertxt.Text) < 1 Then
  Layertxt.Text = 1
End If

If Val(Layertxt.Text) > 4 Then
  Layertxt.Text = 4
End If

Call PlaceGrhCmd_Click

End Sub






Private Sub MapExitTxt_Change()

If Val(MapExitTxt.Text) < 1 Then
  MapExitTxt.Text = 1
End If

If Val(MapExitTxt.Text) > NumMaps Then
  MapExitTxt.Text = NumMaps
End If

Call PlaceExitCmd_Click

End Sub

Private Sub MapLst_DblClick()
'*****************************************************************
'Switch maps
'*****************************************************************

'Check for changes
If MapInfo.Changed = 1 Then
    If MsgBox("Changes have been made to the current map. You will lose all changes if not saved. Save now?", vbYesNo) = vbYes Then
        Call SaveMapData(CurMap)
    End If
End If

'Set user pos and load map
If MapLst.ListIndex <> -1 Then

    'Turn off walkmode
    If WalkMode = True Then
        frmMain.WalkModeChk.value = 0
    End If

    Call SwitchMap(frmMain.MapLst.ItemData(MapLst.ListIndex))
    If MapInfo.StartPos.X > 0 Then
        UserPos.X = MapInfo.StartPos.X
        UserPos.Y = MapInfo.StartPos.Y
    Else
        UserPos.X = WindowTileWidth / 2 + 1
        UserPos.Y = WindowTileHeight / 2 + 1
    End If
    
    EngineRun = True
Else
    MsgBox ("No map selected.")
End If

End Sub



Private Sub MapNameTxt_Change()

MapInfo.Name = MapNameTxt.Text

End Sub

Private Sub MapVersionTxt_Change()

MapInfo.MapVersion = Int(MapVersionTxt.Text)

'Set changed flag
MapInfo.Changed = 1

End Sub

Private Sub MusNumTxt_Change()

MapInfo.Music = MusNumTxt.Text

'Set changed flag
MapInfo.Changed = 1

End Sub

Private Sub NPCLst_Click()

Call PlaceNPCCmd_Click

End Sub

Private Sub OBJAmountTxt_Change()

If Val(OBJAmountTxt.Text) > MAX_INVENORY_OBJS Then
    OBJAmountTxt.Text = 0
End If

If Val(OBJAmountTxt.Text) < 1 Then
    OBJAmountTxt.Text = MAX_INVENORY_OBJS
End If

End Sub

Private Sub ObjLst_Click()

Call PlaceObjCmd_Click

End Sub

Private Sub PlaceBlockCmd_Click()

PlaceGrhCmd.Enabled = True
PlaceBlockCmd.Enabled = False
PlaceExitCmd.Enabled = True
PlaceNPCCmd.Enabled = True
PlaceObjCmd.Enabled = True

End Sub

Private Sub PlaceExitCmd_Click()

PlaceGrhCmd.Enabled = True
PlaceBlockCmd.Enabled = True
PlaceExitCmd.Enabled = False
PlaceNPCCmd.Enabled = True
PlaceObjCmd.Enabled = True

End Sub

Private Sub PlaceGrhCmd_Click()

PlaceGrhCmd.Enabled = False
PlaceBlockCmd.Enabled = True
PlaceExitCmd.Enabled = True
PlaceNPCCmd.Enabled = True
PlaceObjCmd.Enabled = True

End Sub


Private Sub PlaceNPCCmd_Click()

PlaceGrhCmd.Enabled = True
PlaceBlockCmd.Enabled = True
PlaceExitCmd.Enabled = True
PlaceNPCCmd.Enabled = False
PlaceObjCmd.Enabled = True

End Sub


Private Sub PlaceObjCmd_Click()

PlaceGrhCmd.Enabled = True
PlaceBlockCmd.Enabled = True
PlaceExitCmd.Enabled = True
PlaceNPCCmd.Enabled = True
PlaceObjCmd.Enabled = False

End Sub

Private Sub RoomLbl_Click()

End Sub

Private Sub SaveMnu_Click()

If CurMap = 0 Then
    Exit Sub
End If

Call SaveMapData(CurMap)

'Set changed flag
MapInfo.Changed = 0

End Sub


Private Sub SaveNewMnu_Click()

'Add a new map to end of list

If CurMap = 0 Then
    Exit Sub
End If

NumMaps = NumMaps + 1

Call SaveMapData(NumMaps)

'Set changed flag
MapInfo.Changed = 0

RefreshMapList

End Sub


Private Sub WalkModeChk_Click()

ToggleWalkMode

End Sub


Private Sub XExitTxt_Change()

If Val(XExitTxt.Text) < XMinMapSize Then
  XExitTxt.Text = XMinMapSize
End If

If Val(XExitTxt.Text) > XMaxMapSize Then
  XExitTxt.Text = XMaxMapSize
End If

Call PlaceExitCmd_Click

End Sub




Private Sub YExitTxt_Change()

If Val(YExitTxt.Text) < YMinMapSize Then
  YExitTxt.Text = YMinMapSize
End If

If Val(YExitTxt.Text) > YMaxMapSize Then
  YExitTxt.Text = YMaxMapSize
End If

Call PlaceExitCmd_Click

End Sub


