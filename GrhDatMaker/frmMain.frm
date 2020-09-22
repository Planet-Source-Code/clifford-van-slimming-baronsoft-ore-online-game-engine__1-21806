VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grh.dat Maker"
   ClientHeight    =   600
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2640
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   2640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton goCmd 
      Caption         =   "Go!"
      Height          =   405
      Left            =   1530
      TabIndex        =   0
      Top             =   150
      Width           =   1035
   End
   Begin VB.Label statusLbl 
      AutoSize        =   -1  'True
      Caption         =   "Press Go..."
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   240
      Width           =   780
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

IniPath = App.Path & "\"

End Sub

Private Sub goCmd_Click()

'*****************************************************************
'Loads Grh.dat
'*****************************************************************

On Error GoTo ErrorHandler

Dim sX As Integer
Dim sY As Integer
Dim pixelWidth As Integer
Dim pixelHeight As Integer
Dim FileNum As Integer
Dim NumFrames As Integer
Dim Frames(1 To 16) As Integer
Dim Speed As Integer

Dim LastGrh As Integer
Dim TempInt As Integer
Dim Grh As Integer
Dim Frame As Integer
Dim ln As String

LastGrh = Val(GetVar(IniPath & "Grh.ini", "INIT", "NumGrhs"))

'Delete any old file
If FileExist(IniPath & "Grh.dat", vbNormal) = True Then
    Kill IniPath & "Grh.dat"
End If
        
'Open new file
Open IniPath & "Grh.dat" For Binary As #1
Seek #1, 1

'Header
Put #1, , TempInt
Put #1, , TempInt
Put #1, , TempInt
Put #1, , TempInt
Put #1, , TempInt

'Fill Grh List
For Grh = 1 To LastGrh

    statusLbl.Caption = Grh & "/" & LastGrh & " Grhs..."
    statusLbl.Refresh

    'Get line from fisrt file
    ln = GetVar(IniPath & "Grh1.raw", "Graphics", "Grh" & Grh)
    'If not found try othger files
    If ln = "" Then
        ln = GetVar(IniPath & "Grh2.raw", "Graphics", "Grh" & Grh)
    End If
    If ln = "" Then
        ln = GetVar(IniPath & "Grh3.raw", "Graphics", "Grh" & Grh)
    End If
    If ln = "" Then
        ln = GetVar(IniPath & "Grh4.raw", "Graphics", "Grh" & Grh)
    End If
    If ln = "" Then
        ln = GetVar(IniPath & "Grh5.rawt", "Graphics", "Grh" & Grh)
    End If
    
    If ln <> "" Then
    
        'Get number of frames and check
        NumFrames = Val(ReadField(1, ln, 45))
        If NumFrames <= 0 Then GoTo ErrorHandler
    
        'Put grh number
        Put #1, , Grh
        'Put number of frames
        Put #1, , NumFrames
        
        If NumFrames > 1 Then
    
            'Read a animation GRH set
            For Frame = 1 To NumFrames
        
                'Check and put each frame
                Frames(Frame) = Val(ReadField(Frame + 1, ln, 45))
                If Frames(Frame) <= 0 Or Frames(Frame) > LastGrh Then GoTo ErrorHandler
                Put #1, , Frames(Frame)
        
            Next Frame
    
            'Check and put speed
            Speed = Val(ReadField(NumFrames + 2, ln, 45))
            If Speed <= 0 Then GoTo ErrorHandler
            Put #1, , Speed
        
        Else
    
            'check and put normal GRH data
            FileNum = Val(ReadField(2, ln, 45))
            If FileNum <= 0 Then GoTo ErrorHandler
            Put #1, , FileNum
            
            sX = Val(ReadField(3, ln, 45))
            If sX < 0 Then GoTo ErrorHandler
            Put #1, , sX
            
            sY = Val(ReadField(4, ln, 45))
            If sY < 0 Then GoTo ErrorHandler
            Put #1, , sY
            
            pixelWidth = Val(ReadField(5, ln, 45))
            If pixelWidth <= 0 Then GoTo ErrorHandler
            Put #1, , pixelWidth
            
            pixelHeight = Val(ReadField(6, ln, 45))
            If pixelHeight <= 0 Then GoTo ErrorHandler
            Put #1, , pixelHeight

        End If
        
    End If
    
Next Grh
'************************************************

Close #1

statusLbl.Caption = "Done!"
statusLbl.Refresh

Exit Sub

ErrorHandler:
Close #1
MsgBox "Error while loading the GrhX.raw! Stopped at GRH number: " & Grh
statusLbl.Caption = "Error!"
statusLbl.Refresh

End Sub


