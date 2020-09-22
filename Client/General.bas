Attribute VB_Name = "General"
Option Explicit

Sub ReadMapTileStr(TileString As String)
'*****************************************************************
'Takes a tile packet from server, decodes it puts it into the map array
'*****************************************************************
Dim LoopC As Integer
Dim AcumStr As String
Dim TempStr As String
Dim X As Integer, Y As Integer
Dim FieldCounter As Integer

For LoopC = 1 To Len(TileString)
    TempStr = Mid(TileString, LoopC, 1)
    
    If LoopC = Len(TileString) Then
        AcumStr = AcumStr & TempStr
        TempStr = Chr(44)
    End If
    
    If Asc(TempStr) = 44 Then
        Select Case FieldCounter
            Case 0
                X = Val(AcumStr)
            Case 1
                Y = Val(AcumStr)
            Case 2
                MapData(X, Y).Blocked = Val(AcumStr)
            Case Is > 2
                MapData(X, Y).Graphic(Val(Left(AcumStr, 1))).GrhIndex = Val(Right(AcumStr, Len(AcumStr) - 1))
        End Select
        FieldCounter = FieldCounter + 1
        AcumStr = ""
    Else
        AcumStr = AcumStr & TempStr
    End If
    
Next LoopC

If DownloadingMap Then
    frmMain.MainViewLbl = "Downloading New Map: " & Int((Y / YMaxMapSize) * 100) & "%"
End If

End Sub

Sub ClearMapArray()
'*****************************************************************
'Clears all layers
'*****************************************************************

Dim Y As Integer
Dim X As Integer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        'Change blockes status
        MapData(X, Y).Blocked = 0

        'Erase layer 1 and 4
        MapData(X, Y).Graphic(1).GrhIndex = 0
        MapData(X, Y).Graphic(2).GrhIndex = 0
        MapData(X, Y).Graphic(3).GrhIndex = 0
        MapData(X, Y).Graphic(4).GrhIndex = 0

        'Erase characters
        If MapData(X, Y).CharIndex > 0 Then
            Call EraseChar(MapData(X, Y).CharIndex)
        End If

        'Erase Objs
        MapData(X, Y).OBJInfo.OBJIndex = 0
        MapData(X, Y).OBJInfo.Amount = 0
        MapData(X, Y).ObjGrh.GrhIndex = 0

    Next X
Next Y

End Sub
Sub PlayMidi(File As String)
'*****************************************************************
'Plays a Midi using the MCIControl
'*****************************************************************

frmMain.MidiPlayer.Command = "Close"

frmMain.MidiPlayer.FileName = File
    
frmMain.MidiPlayer.Command = "Open"

frmMain.MidiPlayer.Command = "Play"

End Sub




Sub AddtoRichTextBox(RichTextBox As RichTextBox, Text As String, RED As Byte, GREEN As Byte, BLUE As Byte, Bold As Byte, Italic As Byte)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'******************************************

RichTextBox.SelStart = Len(RichTextBox.Text)
RichTextBox.SelLength = 0
RichTextBox.SelColor = RGB(RED, GREEN, BLUE)

If Bold Then
    RichTextBox.SelBold = True
Else
    RichTextBox.SelBold = False
End If

If Italic Then
    RichTextBox.SelItalic = True
Else
    RichTextBox.SelItalic = False
End If

RichTextBox.SelText = Chr(13) & Chr(10) & Text

End Sub
Sub AddtoTextBox(TextBox As TextBox, Text As String)
'******************************************
'Adds text to a text box at the bottom.
'Automatically scrolls to new text.
'******************************************

TextBox.SelStart = Len(TextBox.Text)
TextBox.SelLength = 0


TextBox.SelText = Chr(13) & Chr(10) & Text

End Sub


Sub SaveGameini()
'******************************************
'Saves Game.ini
'******************************************

'update Game.ini
Call WriteVar(IniPath & "Game.ini", "INIT", "Name", UserName)
Call WriteVar(IniPath & "Game.ini", "INIT", "Password", UserPassword)
Call WriteVar(IniPath & "Game.ini", "INIT", "Port", Str(UserPort))

End Sub

Function CheckUserData() As Boolean
'*****************************************************************
'Checks all user data for mistakes and reports them.
'*****************************************************************

Dim LoopC As Integer
Dim CharAscii As Integer

'IP
If UserServerIP = "" Then
    MsgBox ("Server IP box is empty.")
    Exit Function
End If

'Port
If Str(UserPort) = "" Then
    MsgBox ("Port box is empty.")
    Exit Function
End If

'Password
If UserPassword = "" Then
    MsgBox ("Password box is empty.")
    Exit Function
End If

If Len(UserPassword) > 10 Then
    MsgBox ("Password must be 10 characters or less.")
    Exit Function
End If

For LoopC = 1 To Len(UserPassword)

    CharAscii = Asc(Mid$(UserPassword, LoopC, 1))
    If LegalCharacter(CharAscii) = False Then
        MsgBox ("Invalid Password.")
        Exit Function
    End If
    
Next LoopC

'Name
If UserName = "" Then
    MsgBox ("Name box is empty.")
    Exit Function
End If

If Len(UserName) > 30 Then
    MsgBox ("Name must be 30 characters or less.")
    Exit Function
End If

For LoopC = 1 To Len(UserName)

    CharAscii = Asc(Mid$(UserName, LoopC, 1))
    If LegalCharacter(CharAscii) = False Then
        MsgBox ("Invalid Name.")
        Exit Function
    End If
    
Next LoopC

'If all good send true
CheckUserData = True

End Function

Sub UnloadAllForms()
'*****************************************************************
'Unloads all forms
'*****************************************************************

On Error Resume Next

Unload frmConnect
Unload frmMain

End Sub

Function LegalCharacter(KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************

'if backspace allow
If KeyAscii = 8 Then
    LegalCharacter = True
    Exit Function
End If

'Only allow space,numbers,letters and special characters
If KeyAscii < 32 Then
    LegalCharacter = False
    Exit Function
End If

If KeyAscii > 126 Then
    LegalCharacter = False
    Exit Function
End If

'Check for bad special characters in between
If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
    LegalCharacter = False
    Exit Function
End If

'else everything is cool
LegalCharacter = True

End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************

'Set Connected
Connected = True

'Save Game.ini
If frmConnect.SavePassChk.value = 0 Then
    UserPassword = ""
End If
Call SaveGameini

'Unload the connect form
Unload frmConnect

'Load main form
frmMain.Visible = True

End Sub

Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
Static KeyTimer As Integer

'Makes sure keys aren't being pressed to fast
If KeyTimer > 0 Then
    KeyTimer = KeyTimer - 1
    Exit Sub
End If


'Don't allow any these keys during movement..
If UserMoving = 0 Then

    'Move Up
    If GetKeyState(vbKeyUp) < 0 Then
        If LegalPos(UserPos.X, UserPos.Y - 1) Then
            Call SendData("M" & NORTH)
            MoveCharbyHead UserCharIndex, NORTH
            MoveScreen NORTH
        Else
            Call PlayWaveDS(IniPath & "Snd" & "1" & ".wav")
            KeyTimer = 10
        End If
        
        Exit Sub
    End If

    'Move Right
    If GetKeyState(vbKeyRight) < 0 And GetKeyState(vbKeyShift) >= 0 Then

        If LegalPos(UserPos.X + 1, UserPos.Y) Then
            Call SendData("M" & EAST)
            MoveCharbyHead UserCharIndex, EAST
            MoveScreen EAST
        Else
            Call PlayWaveDS(IniPath & "Snd" & "1" & ".wav")
            KeyTimer = 10
        End If

        Exit Sub
    End If

    'Move down
    If GetKeyState(vbKeyDown) < 0 Then
        If LegalPos(UserPos.X, UserPos.Y + 1) Then
            Call SendData("M" & SOUTH)
            MoveCharbyHead UserCharIndex, SOUTH
            MoveScreen SOUTH
        Else
            Call PlayWaveDS(IniPath & "Snd" & "1" & ".wav")
            KeyTimer = 10
        End If

        Exit Sub
    End If

    'Move left
    If GetKeyState(vbKeyLeft) < 0 And GetKeyState(vbKeyShift) >= 0 Then
    
        If LegalPos(UserPos.X - 1, UserPos.Y) Then
            Call SendData("M" & WEST)
            MoveCharbyHead UserCharIndex, WEST
            MoveScreen WEST
        Else
            Call PlayWaveDS(IniPath & "Snd" & "1" & ".wav")
            KeyTimer = 10
        End If

        Exit Sub
    End If

    'Rotate left
    If GetKeyState(vbKeyLeft) < 0 And GetKeyState(vbKeyShift) < 0 Then
    
        Call SendData("<")

        KeyTimer = 10
        Exit Sub
    End If

    'Rotate right
    If GetKeyState(vbKeyRight) < 0 And GetKeyState(vbKeyRight) < 0 Then
    
        Call SendData(">")

        KeyTimer = 10
        Exit Sub
    End If

End If

End Sub




Sub SwitchMap(Map As Integer)
'*****************************************************************
'Loads and switches to a new map
'*****************************************************************

Dim LoopC As Integer
Dim Y As Integer
Dim X As Integer
Dim TempInt As Integer

'Open files
Open App.Path & MapPath & "Map" & Map & ".map" For Binary As #1
Seek #1, 1
        
'map Header
Get #1, , MapInfo.MapVersion
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt
        
'Load arrays
For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        '.dat file
        Get #1, , MapData(X, Y).Blocked
        For LoopC = 1 To 4
            Get #1, , MapData(X, Y).Graphic(LoopC).GrhIndex
            
            'Set up GRH
            If MapData(X, Y).Graphic(LoopC).GrhIndex > 0 Then
                InitGrh MapData(X, Y).Graphic(LoopC), MapData(X, Y).Graphic(LoopC).GrhIndex
            End If
            
        Next LoopC
        'Empty place holders for future expansion
        Get #1, , TempInt
        Get #1, , TempInt
        
        'Erase NPCs
        If MapData(X, Y).CharIndex > 0 Then
            Call EraseChar(MapData(X, Y).CharIndex)
        End If
        
        'Erase OBJs
        MapData(X, Y).OBJInfo.OBJIndex = 0
        MapData(X, Y).OBJInfo.Amount = 0
        MapData(X, Y).ObjGrh.GrhIndex = 0

    Next X
Next Y

Close #1

'Clear out old mapinfo variables
MapInfo.Name = ""
MapInfo.Music = ""

'Set current map
CurMap = Map

End Sub

Public Function ReadField(Pos As Integer, Text As String, SepASCII As Integer) As String
'*****************************************************************
'Gets a field from a string
'*****************************************************************

Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String

Seperator = Chr(SepASCII)
LastPos = 0
FieldNum = 0

For i = 1 To Len(Text)
    CurChar = Mid(Text, i, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = Pos Then
            ReadField = Mid(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next i
FieldNum = FieldNum + 1

If FieldNum = Pos Then
    ReadField = Mid(Text, LastPos + 1)
End If


End Function

Function FileExist(File As String, FileType As VbFileAttribute) As Boolean
'*****************************************************************
'Checks to see if a file exists
'*****************************************************************

If Dir(File, FileType) = "" Then
    FileExist = False
Else
    FileExist = True
End If

End Function

Sub Main()
'*****************************************************************
'Main
'*****************************************************************
Dim LoopC As Integer

'***************************************************
'Start up
'***************************************************
'****** Init vars ******
ENDL = Chr(13) & Chr(10)
ENDC = Chr(1)

'Init Engine
InitTileEngine frmMain.hWnd, 152, 7, 32, 32, 13, 17, 10

'****** Display connect window ******
frmConnect.Visible = True

'****** MidiPlayer INIT ******
frmMain.MidiPlayer.Notify = False
frmMain.MidiPlayer.Wait = False
frmMain.MidiPlayer.Shareable = False
frmMain.MidiPlayer.TimeFormat = mciFormatMilliseconds
frmMain.MidiPlayer.DeviceType = "Sequencer"

'***************************************************
'Main Loop
'***************************************************
prgRun = True
Do While prgRun

    '****** Check Request position timer ******
    If RequestPosTimer > 0 Then
        RequestPosTimer = RequestPosTimer - 1
        If RequestPosTimer = 0 Then
            'Request position Update
            Call SendData("RPU")
        End If
    End If

    '****** Refesh characters on map ******
    Call RefreshAllChars

    '****** Show Next Frame ******

    'Don't draw frame is window is minimized or there is no map loaded
    If frmMain.WindowState <> 1 And CurMap > 0 Then
        
        ShowNextFrame frmMain.Top, frmMain.Left

        '****** Check keys ******
        If DownloadingMap = False Then
            CheckKeys
        End If
    
    End If
            
    '****** Go do other events ******
    DoEvents

Loop
    

'*****************************************************************
'Close Down
'*****************************************************************

'****** Stop any midis ******
mciSendString "close all", 0, 0, 0

'****** Stop Engine ******
DeInitTileEngine

'****** Unload forms and end******
Call UnloadAllForms
End

End Sub

Sub SaveMapData(SaveAs As Integer)
'*****************************************************************
'Saves map data to file
'*****************************************************************

Dim LoopC As Integer
Dim Y As Integer
Dim X As Integer
Dim TempInt As Integer

If FileExist(App.Path & MapPath & "Map" & SaveAs & ".map", vbNormal) = True Then
    Kill App.Path & MapPath & "Map" & SaveAs & ".map"
End If

'Write header info on Map.dat
Call WriteVar(IniPath & "Map.dat", "INIT", "NumMaps", Str(NumMaps))

'Open .map file
Open App.Path & MapPath & "Map" & SaveAs & ".map" For Binary As #1
Seek #1, 1

'map Header
Put #1, , MapInfo.MapVersion
Put #1, , TempInt
Put #1, , TempInt
Put #1, , TempInt
Put #1, , TempInt

'Write .map file
For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        
        '.map file
        Put #1, , MapData(X, Y).Blocked
        For LoopC = 1 To 4
            Put #1, , MapData(X, Y).Graphic(LoopC).GrhIndex
        Next LoopC
        'Empty place holders for future expansion
        Put #1, , TempInt
        Put #1, , TempInt
        
    Next X
Next Y

'Close .map file
Close #1

End Sub

Sub WriteVar(File As String, Main As String, Var As String, value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************

writeprivateprofilestring Main, Var, value, File

End Sub

Function GetVar(File As String, Main As String, Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************

Dim l As Integer
Dim Char As String
Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

szReturn = ""

sSpaces = Space(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish


getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File

GetVar = RTrim(sSpaces)
GetVar = Left(GetVar, Len(GetVar) - 1)

End Function





