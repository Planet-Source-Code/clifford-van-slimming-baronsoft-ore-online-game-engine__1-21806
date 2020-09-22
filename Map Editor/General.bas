Attribute VB_Name = "General"
Option Explicit

'For Get and Write Var
Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'For KeyInput
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Sub RefreshMapList()
'*****************************************************************
'Updates the maps list in the map list
'*****************************************************************
Dim LoopC As Integer
Dim ActualNumMaps As Integer

frmMain.MapLst.Clear

'Add maps to the map list
For LoopC = 1 To NumMaps
    If FileExist(App.Path & MapPath & "Map" & LoopC & ".dat", vbNormal) = True Then
        frmMain.MapLst.AddItem "Map " & LoopC
        frmMain.MapLst.ItemData(frmMain.MapLst.ListCount - 1) = LoopC
        ActualNumMaps = LoopC
    End If
Next LoopC

NumMaps = ActualNumMaps

End Sub


Sub SwitchMap(Map As Integer)
'*****************************************************************
'Loads and switches to a new room
'*****************************************************************
Dim LoopC As Integer
Dim TempInt As Integer
Dim Body As Integer
Dim Head As Integer
Dim Heading As Byte
Dim Y As Integer
Dim X As Integer
   
'Change mouse icon
frmMain.MousePointer = 11
   
'Open files
Open App.Path & MapPath & "Map" & Map & ".map" For Binary As #1
Seek #1, 1
        
Open App.Path & MapPath & "Map" & Map & ".inf" For Binary As #2
Seek #2, 1

'map Header
Get #1, , MapInfo.MapVersion
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt

'inf Header
Get #2, , TempInt
Get #2, , TempInt
Get #2, , TempInt
Get #2, , TempInt
Get #2, , TempInt

'Load arrays
For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        '.map file
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
                                    
                                    
        '.inf file
        
        'Tile exit
        Get #2, , MapData(X, Y).TileExit.Map
        Get #2, , MapData(X, Y).TileExit.X
        Get #2, , MapData(X, Y).TileExit.Y
                      
        'make NPC
        Get #2, , MapData(X, Y).NPCIndex
        If MapData(X, Y).NPCIndex > 0 Then
            Body = Val(GetVar(IniPath & "NPC.dat", "NPC" & MapData(X, Y).NPCIndex, "Body"))
            Head = Val(GetVar(IniPath & "NPC.dat", "NPC" & MapData(X, Y).NPCIndex, "Head"))
            Heading = Val(GetVar(IniPath & "NPC.dat", "NPC" & MapData(X, Y).NPCIndex, "Heading"))
            Call MakeChar(NextOpenChar(), Body, Head, Heading, X, Y)
        End If
        
        'Make obj
        Get #2, , MapData(X, Y).OBJInfo.OBJIndex
        Get #2, , MapData(X, Y).OBJInfo.Amount
        If MapData(X, Y).OBJInfo.OBJIndex > 0 Then
            InitGrh MapData(X, Y).ObjGrh, Val(GetVar(IniPath & "OBJ.dat", "OBJ" & MapData(X, Y).OBJInfo.OBJIndex, "GrhIndex"))
        End If
        
        'Empty place holders for future expansion
        Get #2, , TempInt
        Get #2, , TempInt
             
    Next X
Next Y

'Close files
Close #1
Close #2

'Other Room Data
MapInfo.Name = GetVar(App.Path & MapPath & "Map" & Map & ".dat", "Map" & Map, "Name")
frmMain.MapNameTxt = MapInfo.Name

MapInfo.Music = GetVar(App.Path & MapPath & "Map" & Map & ".dat", "Map" & Map, "MusicNum")
frmMain.MusNumTxt = MapInfo.Music

MapInfo.StartPos.Map = Val(ReadField(1, GetVar(App.Path & MapPath & "Map" & Map & ".dat", "Map" & Map, "StartPos"), 45))
MapInfo.StartPos.X = Val(ReadField(2, GetVar(App.Path & MapPath & "Map" & Map & ".dat", "Map" & Map, "StartPos"), 45))
MapInfo.StartPos.Y = Val(ReadField(3, GetVar(App.Path & MapPath & "Map" & Map & ".dat", "Map" & Map, "StartPos"), 45))
frmMain.StartPosTxt = MapInfo.StartPos.Map & "-" & MapInfo.StartPos.X & "-" & MapInfo.StartPos.Y

frmMain.MapVersionTxt = MapInfo.MapVersion

CurMap = Map
frmMain.MapNameTxt.Text = "Map " & CurMap

'Set changed flag
MapInfo.Changed = 0

'Change mouse icon
frmMain.MousePointer = 0

End Sub
Sub CheckKeys()
'*****************************************************************
'Checks keys
'*****************************************************************

'Check arrow keys
If UserMoving = 0 Then

    If GetKeyState(vbKeyUp) < 0 Then
        If WalkMode = True Then
            If LegalPos(UserPos.X, UserPos.Y - 1) Then
                MoveCharbyHead UserCharIndex, NORTH
                MoveScreen NORTH
            End If
        Else
            MoveScreen NORTH
        End If
        Exit Sub
    End If

    If GetKeyState(vbKeyRight) < 0 Then
        If WalkMode = True Then
            If LegalPos(UserPos.X + 1, UserPos.Y) Then
                MoveCharbyHead UserCharIndex, EAST
                MoveScreen EAST
            End If
        Else
            MoveScreen EAST
        End If
        Exit Sub
    End If

    If GetKeyState(vbKeyDown) < 0 Then
        If WalkMode = True Then
            If LegalPos(UserPos.X, UserPos.Y + 1) Then
                MoveCharbyHead UserCharIndex, SOUTH
                MoveScreen SOUTH
            End If
        Else
            MoveScreen SOUTH
        End If
        Exit Sub
    End If

    If GetKeyState(vbKeyLeft) < 0 Then
        If WalkMode = True Then
            If LegalPos(UserPos.X - 1, UserPos.Y) Then
                MoveCharbyHead UserCharIndex, WEST
                MoveScreen WEST
            End If
        Else
            MoveScreen WEST
        End If
        Exit Sub
    End If

End If

End Sub
Sub ReacttoMouseClick(Button As Integer, tX As Integer, tY As Integer)
'*****************************************************************
'React to mouse button
'*****************************************************************
Dim LoopC As Integer
Dim NPCIndex As Integer
Dim OBJIndex As Integer
Dim Head As Integer
Dim Body As Integer
Dim Heading As Byte

'Right
If Button = vbRightButton Then
    
    'Show Info
    
    'Position
    frmMain.StatTxt.Text = frmMain.StatTxt.Text & ENDL & "Position " & tX & "," & tY & "  Blocked=" & MapData(tX, tY).Blocked
    
    'Exits
    If MapData(tX, tY).TileExit.Map > 0 Then
        frmMain.StatTxt.Text = frmMain.StatTxt.Text & ENDL & "Tile Exit: " & MapData(tX, tY).TileExit.Map & "," & MapData(tX, tY).TileExit.X & "," & MapData(tX, tY).TileExit.Y
    End If
    
    'NPCs
    If MapData(tX, tY).NPCIndex > 0 Then
        frmMain.StatTxt.Text = frmMain.StatTxt.Text & ENDL & "NPC: " & GetVar(IniPath & "NPC.dat", "NPC" & MapData(tX, tY).NPCIndex, "Name")
    End If
    
    'OBJs
    If MapData(tX, tY).OBJInfo.OBJIndex > 0 Then
        frmMain.StatTxt.Text = frmMain.StatTxt.Text & ENDL & "OBJ: " & GetVar(IniPath & "OBJ.dat", "OBJ" & MapData(tX, tY).OBJInfo.OBJIndex, "Name") & "   Amount=" & MapData(tX, tY).OBJInfo.Amount
    End If
    
    'Append
    frmMain.StatTxt.Text = frmMain.StatTxt.Text & ENDL
    frmMain.StatTxt.SelStart = Len(frmMain.StatTxt.Text)
    
    Exit Sub
End If


'Left click
If Button = vbLeftButton Then

    '************** Place grh
    If frmMain.PlaceGrhCmd.Enabled = False Then

        'Erase 2-4
        If frmMain.EraseAllchk.value = 1 Then
            For LoopC = 2 To 4
                MapData(tX, tY).Graphic(LoopC).GrhIndex = 0
            Next LoopC
            Exit Sub
        End If

        'Erase layer
        If frmMain.Erasechk.value = 1 Then
        
            If Val(frmMain.Layertxt.Text) = 1 Then
                MsgBox "Can't Erase Layer 1"
                Exit Sub
            End If
            
            MapData(tX, tY).Graphic(Val(frmMain.Layertxt.Text)).GrhIndex = 0
            Exit Sub
        End If

        'Else Place graphic
        MapData(tX, tY).Blocked = frmMain.Blockedchk.value
        MapData(tX, tY).Graphic(Val(frmMain.Layertxt.Text)).GrhIndex = Val(frmMain.Grhtxt.Text)
        
        'Setup GRH

        InitGrh MapData(tX, tY).Graphic(Val(frmMain.Layertxt.Text)), Val(frmMain.Grhtxt.Text)

    End If
    
    '************** Place blocked tile
    If frmMain.PlaceBlockCmd.Enabled = False Then
        MapData(tX, tY).Blocked = frmMain.Blockedchk.value
    End If

    '************** Place exit
    If frmMain.PlaceExitCmd.Enabled = False Then
        If frmMain.EraseExitChk.value = 0 Then
            MapData(tX, tY).TileExit.Map = Val(frmMain.MapExitTxt.Text)
            MapData(tX, tY).TileExit.X = Val(frmMain.XExitTxt.Text)
            MapData(tX, tY).TileExit.Y = Val(frmMain.YExitTxt.Text)
        Else
            MapData(tX, tY).TileExit.Map = 0
            MapData(tX, tY).TileExit.X = 0
            MapData(tX, tY).TileExit.Y = 0
        End If
    End If

    '************** Place NPC
    If frmMain.PlaceNPCCmd.Enabled = False Then
        If frmMain.EraseNPCChk.value = 0 Then
            If frmMain.NPCLst.ListIndex >= 0 Then
                NPCIndex = frmMain.NPCLst.ListIndex + 1
                Body = Val(GetVar(IniPath & "NPC.dat", "NPC" & NPCIndex, "Body"))
                Head = Val(GetVar(IniPath & "NPC.dat", "NPC" & NPCIndex, "Head"))
                Heading = Val(GetVar(IniPath & "NPC.dat", "NPC" & NPCIndex, "Heading"))
                Call MakeChar(NextOpenChar(), Body, Head, Heading, tX, tY)
                MapData(tX, tY).NPCIndex = NPCIndex
            End If
        Else
            If MapData(tX, tY).NPCIndex > 0 Then
                MapData(tX, tY).NPCIndex = 0
                Call EraseChar(MapData(tX, tY).CharIndex)
            End If
        End If
    End If
    
    '************** Place OBJ
    If frmMain.PlaceObjCmd.Enabled = False Then
        If frmMain.EraseObjChk.value = 0 Then
            If frmMain.ObjLst.ListIndex >= 0 Then
                OBJIndex = frmMain.ObjLst.ListIndex + 1
                InitGrh MapData(tX, tY).ObjGrh, Val(GetVar(IniPath & "OBJ.dat", "OBJ" & OBJIndex, "GrhIndex"))
                MapData(tX, tY).OBJInfo.OBJIndex = OBJIndex
                MapData(tX, tY).OBJInfo.Amount = Val(frmMain.OBJAmountTxt)
            End If
        Else
            MapData(tX, tY).OBJInfo.OBJIndex = 0
            MapData(tX, tY).OBJInfo.Amount = 0
            MapData(tX, tY).ObjGrh.GrhIndex = 0
        End If
    End If
    
    
    'Set changed flag
    MapInfo.Changed = 1
End If

End Sub

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

Sub Main()
'*****************************************************************
'Main
'*****************************************************************
Dim LoopC As Integer

'***************************************************
'Start up
'***************************************************

'****** INIT vars ******
ENDL = Chr(13) & Chr(10)

'Start up engine

'Display form handle, View window offset from 0,0 of display form, Tile Size, Display size in tiles, Screen buffer
If InitTileEngine(frmMain.hWnd, 172, 6, 32, 32, 13, 17, 10) Then

    '****** Load files into memory ******
    Call LoadNPCData
    Call LoadOBJData

End If


'****** Show frmmain ******
frmMain.Show

'****** Refresh map list ******
RefreshMapList

'***************************************************
'Main Loop
'***************************************************
prgRun = True
Do While prgRun

    '****** Show Next Frame ******
    
    'Don't draw frame is window is minimized or there is no map loaded
    If frmMain.WindowState <> 1 And CurMap > 0 Then
        
        ShowNextFrame frmMain.Top, frmMain.Left

        '****** Check keys ******
        CheckKeys
    
    End If
    
    '****** Draw currently selected Grh in ShowPic ******
    If CurrentGrh.GrhIndex = 0 Then
        InitGrh CurrentGrh, 1
    End If
    Call DrawGrhtoHdc(frmMain.ShowPic.hDC, CurrentGrh, 0, 0, 0, 0, SRCCOPY)
    frmMain.ShowPic.Picture = frmMain.ShowPic.Image

    '****** Go do other events ******
    DoEvents

Loop
    

'*****************************************************************
'Close Down
'*****************************************************************

'****** Check if map is saved ******
If MapInfo.Changed = 1 Then
    If MsgBox("Changes have been made to the current map. You will lose all changes if not saved. Save now?", vbYesNo) = vbYes Then
        Call SaveMapData(CurMap)
    End If
End If

'Unload engine
DeInitTileEngine

'****** Unload forms and end******
Unload frmMain
End

End Sub



Function GetVar(File As String, Main As String, Var As String) As String
'*****************************************************************
'Get a var to from a text file
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

Sub WriteVar(File As String, Main As String, Var As String, value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************

writeprivateprofilestring Main, Var, value, File

End Sub

Sub ToggleWalkMode()
'*****************************************************************
'Toggle walk mode on or off
'*****************************************************************

If WalkMode = False Then
    WalkMode = True
Else
    WalkMode = False
End If

If WalkMode = False Then
    'Erase character
    If UserCharIndex > 0 Then
        Call EraseChar(UserCharIndex)
    End If
Else
    'MakeCharacter
    If LegalPos(UserPos.X, UserPos.Y) Then
        Call MakeChar(NextOpenChar(), 1, 1, SOUTH, UserPos.X, UserPos.Y)
        UserCharIndex = MapData(UserPos.X, UserPos.Y).CharIndex
    Else
        MsgBox "Error: Must move over a legal spot first."
        frmMain.WalkModeChk.value = 0
    End If
End If

End Sub

Sub SaveMapData(SaveAs As Integer)
'*****************************************************************
'Save map data to files
'*****************************************************************
Dim LoopC As Integer
Dim TempInt As Integer
Dim Y As Integer
Dim X As Integer

If FileExist(App.Path & MapPath & "Map" & SaveAs & ".dat", vbNormal) = True Then
    If MsgBox("Overwrite existing Map" & SaveAs & ".x files?", vbYesNo) = vbNo Then
        Exit Sub
    End If
End If

'Change mouse icon
frmMain.MousePointer = 11

'Write header info on Map.dat
Call WriteVar(IniPath & "Map.dat", "INIT", "NumMaps", Str(NumMaps))

'Erase old files if the exist
If FileExist(App.Path & MapPath & "Map" & SaveAs & ".map", vbNormal) = True Then
    Kill App.Path & MapPath & "Map" & SaveAs & ".map"
End If

If FileExist(App.Path & MapPath & "Map" & SaveAs & ".inf", vbNormal) = True Then
    Kill App.Path & MapPath & "Map" & SaveAs & ".inf"
End If

'Open .map file
Open App.Path & MapPath & "Map" & SaveAs & ".map" For Binary As #1
Seek #1, 1

'Open .inf file
Open App.Path & MapPath & "Map" & SaveAs & ".inf" For Binary As #2
Seek #2, 1

'map Header
Put #1, , MapInfo.MapVersion
Put #1, , TempInt
Put #1, , TempInt
Put #1, , TempInt
Put #1, , TempInt

'inf Header
Put #2, , TempInt
Put #2, , TempInt
Put #2, , TempInt
Put #2, , TempInt
Put #2, , TempInt

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
        
        
        '.inf file
        'Tile exit
        Put #2, , MapData(X, Y).TileExit.Map
        Put #2, , MapData(X, Y).TileExit.X
        Put #2, , MapData(X, Y).TileExit.Y
        
        'NPC
        Put #2, , MapData(X, Y).NPCIndex
        
        'Object
        Put #2, , MapData(X, Y).OBJInfo.OBJIndex
        Put #2, , MapData(X, Y).OBJInfo.Amount
        
        'Empty place holders for future expansion
        Put #2, , TempInt
        Put #2, , TempInt
        
    Next X
Next Y

'Close .map file
Close #1

'Close .inf file
Close #2

'write .dat file
Call WriteVar(App.Path & MapPath & "Map" & SaveAs & ".dat", "Map" & SaveAs, "Name", MapInfo.Name)
Call WriteVar(App.Path & MapPath & "Map" & SaveAs & ".dat", "Map" & SaveAs, "MusicNum", MapInfo.Music)
Call WriteVar(App.Path & MapPath & "Map" & SaveAs & ".dat", "Map" & SaveAs, "StartPos", MapInfo.StartPos.Map & "-" & MapInfo.StartPos.X & "-" & MapInfo.StartPos.Y)

'Change mouse icon
frmMain.MousePointer = 0

MsgBox ("Current map saved as map # " & SaveAs)

End Sub

Sub LoadOBJData()
'*****************************************************************
'Setup OBJ list
'*****************************************************************
Dim NumOBJs As Integer
Dim Obj As Integer

'Get Number of Maps
NumOBJs = Val(GetVar(IniPath & "OBJ.dat", "INIT", "NumOBJs"))

'Add OBJs to the OBJ list
For Obj = 1 To NumOBJs
    frmMain.ObjLst.AddItem GetVar(IniPath & "OBJ.dat", "OBJ" & Obj, "Name")
Next Obj

End Sub

Sub LoadNPCData()
'*****************************************************************
'Setup NPC list
'*****************************************************************
Dim NumNPCs As Integer
Dim NPC As Integer

'Get Number of Maps
NumNPCs = Val(GetVar(IniPath & "NPC.dat", "INIT", "NumNPCs"))

'Add NPCs to the NPC list
For NPC = 1 To NumNPCs
    frmMain.NPCLst.AddItem GetVar(IniPath & "NPC.dat", "NPC" & NPC, "Name")
Next NPC

End Sub

