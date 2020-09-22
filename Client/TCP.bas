Attribute VB_Name = "TCP"
Option Explicit

Sub HandleData(rData As String)
'*********************************************
'Handle all data from server
'*********************************************
Dim retVal As Variant
Dim X As Integer
Dim Y As Integer
Dim CharIndex As Integer
Dim ServerHandle As Integer
Dim TempInt As Integer
Dim TempStr As String
Dim Slot As Integer

'**************** Communication stuff ****************

'Send to Rectxt
If Left$(rData, 1) = "@" Then
    rData = Right$(rData, Len(rData) - 1)
    
    AddtoRichTextBox frmMain.RecTxt, ReadField(1, rData, 126), Val(ReadField(2, rData, 126)), Val(ReadField(3, rData, 126)), Val(ReadField(4, rData, 126)), Val(ReadField(5, rData, 126)), Val(ReadField(6, rData, 126))
    
    Exit Sub
End If

'Urgant MsgBox
If Left$(rData, 2) = "!!" Then
    rData = Right$(rData, Len(rData) - 2)
    MsgBox rData, vbApplicationModal
    Exit Sub
End If

'MsgBox
If Left$(rData, 1) = "!" Then
    rData = Right$(rData, Len(rData) - 1)
    MsgBox rData
    Exit Sub
End If

'**************** Intitialization stuff ****************

'Get UserServerIndex
If Left$(rData, 3) = "SUI" Then
    rData = Right$(rData, Len(rData) - 3)
    UserIndex = (Val(rData))
    Exit Sub
End If

'Get UserCharIndex
If Left$(rData, 3) = "SUC" Then
    rData = Right$(rData, Len(rData) - 3)
    UserCharIndex = (Val(rData))
    UserPos = CharList(UserCharIndex).Pos
    Exit Sub
End If

'Set user's screen pos
If Left$(rData, 3) = "SSP" Then
    rData = Right$(rData, Len(rData) - 3)
    UserPos.X = ReadField(1, rData, 44)
    UserPos.Y = ReadField(2, rData, 44)
    Exit Sub
End If

'Set user position
If Left$(rData, 3) = "SUP" Then
    rData = Right$(rData, Len(rData) - 3)
    
    X = ReadField(1, rData, 44)
    Y = ReadField(2, rData, 44)
    
    MapData(UserPos.X, UserPos.Y).CharIndex = 0
    MapData(X, Y).CharIndex = UserCharIndex
    
    UserPos.X = X
    UserPos.Y = Y
    CharList(UserCharIndex).Pos = UserPos
    
    Exit Sub
End If

'**************** Map stuff ****************

'Load map
If Left$(rData, 3) = "SCM" Then
    rData = Right$(rData, Len(rData) - 3)
    
    'Stop engine
    EngineRun = False
    
    'Set switching map flag
    DownloadingMap = True
    
    'Get Version Num
    If FileExist(App.Path & MapPath & "Map" & ReadField(1, rData, 44) & ".map", vbNormal) Then
        Open App.Path & MapPath & "Map" & ReadField(1, rData, 44) & ".map" For Binary As #1
        Seek #1, 1
        Get #1, , TempInt
        Close #1
        If TempInt = Val(ReadField(2, rData, 44)) Then
            'Correct Version
            SwitchMap ReadField(1, rData, 44)
            SendData "DLM" 'Tell the server we are done loading map
        Else
            'Not correct version
            SendData "RMU" & ReadField(1, rData, 44)
        End If
    Else
        'Didn't find map
        SendData "RMU" & ReadField(1, rData, 44)
    End If
    
    Exit Sub
End If

'Start Map Transfer
If Left$(rData, 3) = "SMT" Then
    rData = Right$(rData, Len(rData) - 3)
    
    MapInfo.MapVersion = Val(rData)
    
    ClearMapArray
    frmMain.MainViewLbl.Visible = True
    
    Exit Sub
End If

'Set Map Tile
If Left$(rData, 3) = "CMT" Then
    rData = Right$(rData, Len(rData) - 3)
    
    ReadMapTileStr rData
    
    SendData "RNT"
    
    Exit Sub
    
End If

'End Map Transfer
If Left$(rData, 3) = "EMT" Then
    rData = Right$(rData, Len(rData) - 3)
    If Val(rData) > NumMaps Then NumMaps = Val(rData)
    SaveMapData Val(rData)
    SwitchMap Val(rData)
    frmMain.MainViewLbl.Visible = False
    SendData "DLM" 'Tell the server we are done loading map
End If

'Done switching maps
If rData = "DSM" Then
    DownloadingMap = False
    EngineRun = True
    Exit Sub
End If

'Change map name
If Left$(rData, 3) = "SMN" Then
    MapInfo.Name = Right$(rData, Len(rData) - 3)
    frmMain.MapNameLbl.Caption = MapInfo.Name
    Exit Sub
End If

'**************** Character and object stuff ****************

'Ignore this stuff if downloading a map
If DownloadingMap = False Then


    'Make Char
    If Left$(rData, 3) = "MAC" Then
        rData = Right$(rData, Len(rData) - 3)
    
        CharIndex = ReadField(4, rData, 44)
        X = ReadField(5, rData, 44)
        Y = ReadField(6, rData, 44)
    
        Call MakeChar(CharIndex, ReadField(1, rData, 44), ReadField(2, rData, 44), ReadField(3, rData, 44), X, Y)

        Exit Sub
    End If

    'Erase Char
    If Left$(rData, 3) = "ERC" Then
        rData = Right$(rData, Len(rData) - 3)

        Call EraseChar(Val(rData))

        Exit Sub
    End If

    'Move Char
    If Left$(rData, 3) = "MOC" Then
        rData = Right$(rData, Len(rData) - 3)

        CharIndex = Val(ReadField(1, rData, 44))

        Call MoveCharbyPos(CharIndex, ReadField(2, rData, 44), ReadField(3, rData, 44))

        Exit Sub
    End If

    'Change Char
    If Left$(rData, 3) = "CHC" Then
        rData = Right$(rData, Len(rData) - 3)

        CharIndex = Val(ReadField(1, rData, 44))

        CharList(CharIndex).Body = BodyData(Val(ReadField(2, rData, 44)))
        CharList(CharIndex).Head = HeadData(Val(ReadField(3, rData, 44)))
        CharList(CharIndex).Heading = Val(ReadField(4, rData, 44))

        Exit Sub
    End If

    'Make Obj layer
    If Left$(rData, 3) = "MOB" Then
        rData = Right$(rData, Len(rData) - 3)
        X = Val(ReadField(2, rData, 44))
        Y = Val(ReadField(3, rData, 44))
        MapData(X, Y).ObjGrh.GrhIndex = Val(ReadField(1, rData, 44))
        InitGrh MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex
        Exit Sub
    End If

    'Erase Obj layer
    If Left$(rData, 3) = "EOB" Then
        rData = Right$(rData, Len(rData) - 3)
        X = Val(ReadField(1, rData, 44))
        Y = Val(ReadField(2, rData, 44))
        MapData(X, Y).ObjGrh.GrhIndex = 0
        Exit Sub
    End If


End If

'**************** Status stuff ****************

'Update Main Stats
If Left$(rData, 3) = "SST" Then
    rData = Right$(rData, Len(rData) - 3)

    UserMaxHP = Val(ReadField(1, rData, 44))
    UserMinHP = Val(ReadField(2, rData, 44))
    UserMaxMAN = Val(ReadField(3, rData, 44))
    UserMinMAN = Val(ReadField(4, rData, 44))
    UserMaxSTA = Val(ReadField(5, rData, 44))
    UserMinSTA = Val(ReadField(6, rData, 44))
    UserGLD = Val(ReadField(7, rData, 44))
    UserLvl = Val(ReadField(8, rData, 44))

    frmMain.HPShp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 150)

    frmMain.MANShp.Width = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 150)

    frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 150)

    frmMain.GldLbl.Caption = UserGLD
    frmMain.LvlLbl.Caption = UserLvl

    Exit Sub
End If

'Set Inventory Slot
If Left$(rData, 3) = "SIS" Then
    rData = Right$(rData, Len(rData) - 3)

    Slot = ReadField(1, rData, 44)
    UserInventory(Slot).OBJIndex = ReadField(2, rData, 44)
    UserInventory(Slot).Name = ReadField(3, rData, 44)
    UserInventory(Slot).Amount = ReadField(4, rData, 44)
    UserInventory(Slot).Equipped = ReadField(5, rData, 44)
    UserInventory(Slot).GrhIndex = Val(ReadField(6, rData, 44))
    
    TempStr = ""
    If UserInventory(Slot).Equipped = 1 Then
        TempStr = TempStr & "(Eqp)"
    End If
    
    If UserInventory(Slot).Amount > 0 Then
        TempStr = TempStr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).Name
    Else
        TempStr = TempStr & UserInventory(Slot).Name
    End If
    
    frmMain.ObjLst.List(Slot - 1) = TempStr
    
    Exit Sub
End If

'**************** Sound stuff ****************

'Play midi
If Left$(rData, 3) = "PLM" Then
    rData = Right$(rData, Len(rData) - 3)
    
    CurMidi = IniPath & "Mus" & Val(ReadField(1, rData, 45)) & ".mid"
    LoopMidi = Val(ReadField(2, rData, 45))
    Call PlayMidi(CurMidi)

    
    Exit Sub
End If

'Play Wave
If Left$(rData, 3) = "PLW" Then
    rData = Right$(rData, Len(rData) - 3)
    
        Call PlayWaveDS(IniPath & "Snd" & rData & ".wav")
        
    Exit Sub
End If

End Sub

Sub SendData(sdData As String)
'*********************************************
'Attach a ENDC to a string and send to server
'*********************************************
Dim retcode

sdData = sdData & ENDC

'To avoid spam set a limit
If Len(sdData) > 300 Then
    Exit Sub
End If

retcode = frmMain.Socket1.Write(sdData, Len(sdData))

End Sub

Sub Login()
'*********************************************
'Send login strings
'*********************************************

'Pre-saved character
If SendNewChar = False Then
    SendData ("LOGIN" & UserName & "," & UserPassword & "," & App.Major & "." & App.Minor & "." & App.Revision)
End If

'New character
If SendNewChar = True Then
    SendData ("NLOGIN" & UserName & "," & UserPassword & "," & UserBody & "," & UserHead & "," & App.Major & "." & App.Minor & "." & App.Revision)
End If

End Sub


