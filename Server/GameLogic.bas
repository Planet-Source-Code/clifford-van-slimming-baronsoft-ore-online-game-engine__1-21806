Attribute VB_Name = "GameLogic"
Option Explicit
Sub SendNextMapTile(ByVal UserIndex As Integer)
'*****************************************************************
'Send a map tile to a user
'*****************************************************************
Dim LoopC As Integer
Dim ln As String
Dim TempInt As Integer
              
If UserList(UserIndex).Counters.SendMapCounter.Y > YMaxMapSize Then
    SendData ToIndex, UserIndex, 0, "EMT" & UserList(UserIndex).Counters.SendMapCounter.Map
    UserList(UserIndex).Flags.DownloadingMap = 0
    UserList(UserIndex).Counters.SendMapCounter.X = 0
    UserList(UserIndex).Counters.SendMapCounter.Y = 0
    UserList(UserIndex).Counters.SendMapCounter.Map = 0
Else
    
    ln = UserList(UserIndex).Counters.SendMapCounter.X & "," & UserList(UserIndex).Counters.SendMapCounter.Y & "," & MapData(UserList(UserIndex).Counters.SendMapCounter.Map, UserList(UserIndex).Counters.SendMapCounter.X, UserList(UserIndex).Counters.SendMapCounter.Y).Blocked
    For LoopC = 1 To 4
        TempInt = MapData(UserList(UserIndex).Counters.SendMapCounter.Map, UserList(UserIndex).Counters.SendMapCounter.X, UserList(UserIndex).Counters.SendMapCounter.Y).Graphic(LoopC)
        If TempInt > 0 Then
            ln = ln & "," & LoopC & TempInt
        End If
    Next LoopC
                
    SendData ToIndex, UserIndex, 0, "CMT" & ln
    
    UserList(UserIndex).Counters.SendMapCounter.X = UserList(UserIndex).Counters.SendMapCounter.X + 1
    If UserList(UserIndex).Counters.SendMapCounter.X > XMaxMapSize Then
        UserList(UserIndex).Counters.SendMapCounter.X = XMinMapSize
        UserList(UserIndex).Counters.SendMapCounter.Y = UserList(UserIndex).Counters.SendMapCounter.Y + 1
    End If
    
    UserList(UserIndex).Flags.ReadyForNextTile = 0
    
End If

End Sub

Sub ChangeUserChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, UserIndex As Integer, Body As Integer, Head As Integer, Heading As Byte)
'*****************************************************************
'Changes a user char's head,body and heading
'*****************************************************************

UserList(UserIndex).Char.Body = Body
UserList(UserIndex).Char.Head = Head
UserList(UserIndex).Char.Heading = Heading

Call SendData(sndRoute, sndIndex, sndMap, "CHC" & UserList(UserIndex).Char.CharIndex & "," & Body & "," & Head & "," & Heading)

End Sub

Sub ChangeNPCChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, NPCIndex As Integer, Body As Integer, Head As Integer, Heading As Byte)
'*****************************************************************
'Changes a NPC char's head,body and heading
'*****************************************************************

NPCList(NPCIndex).Char.Body = Body
NPCList(NPCIndex).Char.Head = Head
NPCList(NPCIndex).Char.Heading = Heading

Call SendData(sndRoute, sndIndex, sndMap, "CHC" & NPCList(NPCIndex).Char.CharIndex & "," & Body & "," & Head & "," & Heading)

End Sub

Sub CheckUserLevel(ByVal UserIndex As Integer)
'*****************************************************************
'Checks user's exp and levels user up
'*****************************************************************

'Make sure user hasn't reached max level
If UserList(UserIndex).Stats.ELV = STAT_MAXELV Then
    UserList(UserIndex).Stats.EXP = 0
    UserList(UserIndex).Stats.ELU = 0
    Exit Sub
End If

'If exp >= then elu then level up user
If UserList(UserIndex).Stats.EXP >= UserList(UserIndex).Stats.ELU Then

    UserList(UserIndex).Stats.ELV = UserList(UserIndex).Stats.ELV + 1
    UserList(UserIndex).Stats.EXP = 0
    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.5

    AddtoVar UserList(UserIndex).Stats.MaxHP, 2, STAT_MAXHP
    AddtoVar UserList(UserIndex).Stats.MaxSTA, 2, STAT_MAXSTA
    AddtoVar UserList(UserIndex).Stats.MaxMAN, 2, STAT_MAXMAN
    
    AddtoVar UserList(UserIndex).Stats.MaxHIT, 2, STAT_MAXHIT
    AddtoVar UserList(UserIndex).Stats.MinHIT, 2, STAT_MAXHIT
    AddtoVar UserList(UserIndex).Stats.DEF, 2, STAT_MAXDEF
    
    AddtoVar UserList(UserIndex).Stats.MET, 1, STAT_MAXSTAT
    AddtoVar UserList(UserIndex).Stats.FIT, 1, STAT_MAXSTAT
    
    SendData ToIndex, UserIndex, 0, "@You level up!" & FONTTYPE_INFO
    SendUserStatsBox UserIndex

End If

End Sub

Sub DoTileEvents(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'*****************************************************************
'Do any events on a tile
'*****************************************************************

'Check for tile exit
If MapData(Map, X, Y).TileExit.Map > 0 Then
    If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y) Then
        Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y)
    End If
End If

End Sub

Function InMapBounds(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************

If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function

Sub KillNPC(ByVal NPCIndex As Integer)
'*****************************************************************
'Kill a NPC
'*****************************************************************

'Set health back to 100%
NPCList(NPCIndex).Stats.MinHP = NPCList(NPCIndex).Stats.MaxHP

'Erase it from map
EraseNPCChar ToMap, 0, NPCList(NPCIndex).Pos.Map, NPCIndex

'Set respawn wait
NPCList(NPCIndex).Counters.RespawnCounter = NPCList(NPCIndex).RespawnWait

End Sub

Sub SpawnNPC(ByVal NPCIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'*****************************************************************
'Places a NPC that has been Opened
'*****************************************************************

Dim TempPos As WorldPos

NPCList(NPCIndex).Pos.Map = Map
NPCList(NPCIndex).Pos.X = X
NPCList(NPCIndex).Pos.Y = Y

'Find a place to put npc
Call ClosestLegalPos(NPCList(NPCIndex).Pos, TempPos)
If LegalPos(TempPos.Map, TempPos.X, TempPos.Y) = False Then
    Exit Sub
End If

'Set vars
NPCList(NPCIndex).Pos = TempPos
NPCList(NPCIndex).StartPos = TempPos

'Make NPC Char
Call MakeNPCChar(ToMap, 0, TempPos.Map, NPCIndex, TempPos.Map, TempPos.X, TempPos.Y)

End Sub

Sub KillUser(ByVal UserIndex As Integer)
'*****************************************************************
'Kill a user
'*****************************************************************
Dim TempPos As WorldPos

'Set user health back to full
UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP

'Find a place to put user
Call ClosestLegalPos(ResPos, TempPos)
If LegalPos(TempPos.Map, TempPos.X, TempPos.Y) = False Then
    Call SendData(ToIndex, UserIndex, 0, "!!No legal position found: Please try again.")
    CloseUser (UserIndex)
    Exit Sub
End If

'Warp him there
WarpUserChar UserIndex, TempPos.Map, TempPos.X, TempPos.Y

End Sub

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'*****************************************************************
'Use/Equip a inventory item
'*****************************************************************
Dim Obj As ObjData

Obj = ObjData(UserList(UserIndex).Object(Slot).ObjIndex)

Select Case Obj.ObjType

    Case OBJTYPE_USEONCE
    
        'use item
        AddtoVar UserList(UserIndex).Stats.MaxHP, Obj.MaxHP, 999
        AddtoVar UserList(UserIndex).Stats.MinHP, Obj.MinHP, UserList(UserIndex).Stats.MaxHP
        
        'Remove from inventory
        UserList(UserIndex).Object(Slot).Amount = UserList(UserIndex).Object(Slot).Amount - 1
        If UserList(UserIndex).Object(Slot).Amount <= 0 Then
            UserList(UserIndex).Object(Slot).ObjIndex = 0
        End If


    Case OBJTYPE_WEAPON
        
        'If currently equipped remove instead
        If UserList(UserIndex).Object(Slot).Equipped Then
            RemoveInvItem UserIndex, Slot
            Exit Sub
        End If
        
        'Remove old item if exists
        If UserList(UserIndex).WeaponEqpObjIndex > 0 Then
            RemoveInvItem UserIndex, UserList(UserIndex).WeaponEqpSlot
        End If

        'Equip
        UserList(UserIndex).Stats.MaxHIT = UserList(UserIndex).Stats.MaxHIT + Obj.MaxHIT
        UserList(UserIndex).Stats.MinHIT = UserList(UserIndex).Stats.MinHIT + Obj.MinHIT
        UserList(UserIndex).Object(Slot).Equipped = 1
        UserList(UserIndex).WeaponEqpObjIndex = UserList(UserIndex).Object(Slot).ObjIndex
        UserList(UserIndex).WeaponEqpSlot = Slot

    Case OBJTYPE_ARMOUR

        'If currently equipped remove instead
        If UserList(UserIndex).Object(Slot).Equipped Then
            RemoveInvItem UserIndex, Slot
            Exit Sub
        End If

        'Remove old item if exists
        If UserList(UserIndex).ArmourEqpObjIndex > 0 Then
            RemoveInvItem UserIndex, UserList(UserIndex).ArmourEqpSlot
        End If

        'Equip
        UserList(UserIndex).Stats.DEF = UserList(UserIndex).Stats.DEF + Obj.DEF
        UserList(UserIndex).Object(Slot).Equipped = 1
        UserList(UserIndex).ArmourEqpObjIndex = UserList(UserIndex).Object(Slot).ObjIndex
        UserList(UserIndex).ArmourEqpSlot = Slot

End Select

'Update user's stats and inventory
SendUserStatsBox UserIndex
UpdateUserInv True, UserIndex, 0

End Sub

Sub AddtoVar(Var As Variant, ByVal Addon As Variant, ByVal Max As Variant)
'*****************************************************************
'Adds a value to a variable respecting a max value
'*****************************************************************

If Var >= Max Then
    Var = Max
    Exit Sub
End If

Var = Var + Addon
If Var > Max Then
    Var = Max
End If

End Sub

Sub RemoveInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'*****************************************************************
'Unequip a inventory item
'*****************************************************************

Dim Obj As ObjData

Obj = ObjData(UserList(UserIndex).Object(Slot).ObjIndex)


Select Case Obj.ObjType


    Case OBJTYPE_WEAPON

        UserList(UserIndex).Stats.MaxHIT = UserList(UserIndex).Stats.MaxHIT - Obj.MaxHIT
        UserList(UserIndex).Stats.MinHIT = UserList(UserIndex).Stats.MinHIT - Obj.MinHIT

        UserList(UserIndex).Object(Slot).Equipped = 0
        UserList(UserIndex).WeaponEqpObjIndex = 0
        UserList(UserIndex).WeaponEqpSlot = 0

    Case OBJTYPE_ARMOUR

        UserList(UserIndex).Stats.DEF = UserList(UserIndex).Stats.DEF - Obj.DEF

        UserList(UserIndex).Object(Slot).Equipped = 0
        UserList(UserIndex).ArmourEqpObjIndex = 0
        UserList(UserIndex).ArmourEqpSlot = 0

End Select

SendUserStatsBox UserIndex
UpdateUserInv True, UserIndex, 0

End Sub

Function NextOpenCharIndex() As Integer
'*****************************************************************
'Finds the next open CharIndex in Charlist
'*****************************************************************
Dim LoopC As Integer

For LoopC = 1 To LastChar + 1
    If CharList(LoopC) = 0 Then
        NextOpenCharIndex = LoopC
        NumChars = NumChars + 1
        If LoopC > LastChar Then LastChar = LoopC
        Exit Function
    End If
Next LoopC

End Function

Function NextOpenUser() As Integer
'*****************************************************************
'Finds the next open UserIndex in UserList
'*****************************************************************
Dim LoopC As Integer
  
LoopC = 1
  
Do Until UserList(LoopC).Flags.UserLogged = 0
    LoopC = LoopC + 1
Loop
  
NextOpenUser = LoopC

End Function

Function NextOpenNPC() As Integer
'*****************************************************************
'Finds the next open UserIndex in UserList
'*****************************************************************
Dim LoopC As Integer
  
LoopC = 1
  
Do Until NPCList(LoopC).Flags.NPCActive = 0
    LoopC = LoopC + 1
Loop
  
NextOpenNPC = LoopC

End Function

Sub ClosestLegalPos(Pos As WorldPos, nPos As WorldPos)
'*****************************************************************
'Finds the closest legal tile to Pos and stores it in nPos
'*****************************************************************
Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = Pos.Map

Do While LegalPos(Pos.Map, nPos.X, nPos.Y) = False
    
    If LoopC > 10 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.Y - LoopC To Pos.Y + LoopC
        For tX = Pos.X - LoopC To Pos.X + LoopC
        
            If LegalPos(nPos.Map, tX, tY) = True Then
                nPos.X = tX
                nPos.Y = tY
                tX = Pos.X + LoopC
                tY = Pos.Y + LoopC
            End If
        
        Next tX
    Next tY
    
    LoopC = LoopC + 1
    
Loop

If Notfound = True Then
    nPos.X = 0
    nPos.Y = 0
End If

End Sub

Function NameIndex(ByVal Name As String) As Integer
'*****************************************************************
'Searches userlist for a name and return userindex
'*****************************************************************
Dim UserIndex As Integer
  
'check for bad name
If Name = "" Then
    NameIndex = 0
    Exit Function
End If
  
UserIndex = 1
Do Until UCase(Left$(UserList(UserIndex).Name, Len(Name))) = UCase(Name)
    
    UserIndex = UserIndex + 1
    
    If UserIndex > LastUser Then
        UserIndex = 0
        Exit Do
    End If
    
Loop
  
NameIndex = UserIndex

End Function

Sub NPCAI(ByVal NPCIndex As Integer)
'*****************************************************************
'Moves NPC based on it's .movement value
'*****************************************************************
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer

'Look for someone to attack if hostile
If NPCList(NPCIndex).Hostile Then

    'Check in all directions
    For HeadingLoop = NORTH To WEST
        nPos = NPCList(NPCIndex).Pos
        HeadtoPos HeadingLoop, nPos
        
        'if a legal pos and a user is found attack
        If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
            If MapData(nPos.Map, nPos.X, nPos.Y).UserIndex > 0 Then
                'Face NPC to target
                ChangeNPCChar ToMap, 0, nPos.Map, NPCIndex, NPCList(NPCIndex).Char.Body, NPCList(NPCIndex).Char.Head, HeadingLoop
                'Attack
                NPCAttackUser NPCIndex, MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
                'Don't move if fighting
                Exit Sub
            End If
        End If
        
    Next HeadingLoop
End If


'Movement
Select Case NPCList(NPCIndex).Movement

    'Stand
    Case 1
        'Do nothing
        
    'Move randomly
    Case 2
        Call MoveNPCChar(NPCIndex, Int(RandomNumber(1, 4)))

    'Go towards any nearby Users
    Case 3
        For Y = NPCList(NPCIndex).Pos.Y - 5 To NPCList(NPCIndex).Pos.Y + 5    'Makes a loop that looks at
            For X = NPCList(NPCIndex).Pos.X - 5 To NPCList(NPCIndex).Pos.X + 5   '5 tiles in every direction

                'Make sure tile is legal
                If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then
                
                    'look for a user
                    If MapData(NPCList(NPCIndex).Pos.Map, X, Y).UserIndex > 0 Then
                        'Move towards user
                        tHeading = FindDirection(NPCList(NPCIndex).Pos, UserList(MapData(NPCList(NPCIndex).Pos.Map, X, Y).UserIndex).Pos)
                        MoveNPCChar NPCIndex, tHeading
                        'Leave sub
                        Exit Sub
                    End If
                    
                End If
                     
            Next X
        Next Y

End Select

End Sub

Function OpenNPC(ByVal NPCNumber As Integer) As Integer
'*****************************************************************
'Loads a NPC and returns its index
'*****************************************************************
Dim NPCIndex As Integer
Dim NPCFile As String

'Set NPC file
NPCFile = IniPath & "NPC.dat"

'Find next open NPCindex
NPCIndex = NextOpenNPC

'Load stats from file
NPCList(NPCIndex).Name = GetVar(NPCFile, "NPC" & NPCNumber, "Name")
NPCList(NPCIndex).Desc = GetVar(NPCFile, "NPC" & NPCNumber, "Desc")
NPCList(NPCIndex).Movement = Val(GetVar(NPCFile, "NPC" & NPCNumber, "Movement"))
NPCList(NPCIndex).RespawnWait = Val(GetVar(NPCFile, "NPC" & NPCNumber, "RespawnWait"))

NPCList(NPCIndex).Char.Body = Val(GetVar(NPCFile, "NPC" & NPCNumber, "Body"))
NPCList(NPCIndex).Char.Head = Val(GetVar(NPCFile, "NPC" & NPCNumber, "Head"))
NPCList(NPCIndex).Char.Heading = Val(GetVar(NPCFile, "NPC" & NPCNumber, "Heading"))

NPCList(NPCIndex).Attackable = Val(GetVar(NPCFile, "NPC" & NPCNumber, "Attackable"))
NPCList(NPCIndex).Hostile = Val(GetVar(NPCFile, "NPC" & NPCNumber, "Hostile"))
NPCList(NPCIndex).GiveEXP = Val(GetVar(NPCFile, "NPC" & NPCNumber, "GiveEXP"))
NPCList(NPCIndex).GiveGLD = Val(GetVar(NPCFile, "NPC" & NPCNumber, "GiveGLD"))

NPCList(NPCIndex).Stats.MaxHP = Val(GetVar(NPCFile, "NPC" & NPCNumber, "MaxHP"))
NPCList(NPCIndex).Stats.MinHP = Val(GetVar(NPCFile, "NPC" & NPCNumber, "MinHP"))
NPCList(NPCIndex).Stats.MaxHIT = Val(GetVar(NPCFile, "NPC" & NPCNumber, "MaxHIT"))
NPCList(NPCIndex).Stats.MinHIT = Val(GetVar(NPCFile, "NPC" & NPCNumber, "MinHIT"))
NPCList(NPCIndex).Stats.DEF = Val(GetVar(NPCFile, "NPC" & NPCNumber, "DEF"))

'Setup NPC
NPCList(NPCIndex).Flags.NPCActive = 1

'Update NPC counters
If NPCIndex > LastNPC Then LastNPC = NPCIndex
NumNPCs = NumNPCs + 1

'Return new NPCIndex
OpenNPC = NPCIndex

End Function

Sub EraseObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal Num As Integer, ByVal Map As Byte, ByVal X As Integer, ByVal Y As Integer)
'*****************************************************************
'Erase a object
'*****************************************************************

MapData(Map, X, Y).ObjInfo.Amount = MapData(Map, X, Y).ObjInfo.Amount - Num

If MapData(Map, X, Y).ObjInfo.Amount <= 0 Then
    MapData(Map, X, Y).ObjInfo.ObjIndex = 0
    MapData(Map, X, Y).ObjInfo.Amount = 0
    Call SendData(sndRoute, sndIndex, sndMap, "EOB" & X & "," & Y)
End If

End Sub

Sub MakeObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Obj As Obj, Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'*****************************************************************
'Erase a object
'*****************************************************************

MapData(Map, X, Y).ObjInfo = Obj
Call SendData(sndRoute, sndIndex, sndMap, "MOB" & ObjData(Obj.ObjIndex).GRHIndex & "," & X & "," & Y)

End Sub

Sub GetObj(ByVal UserIndex As Integer)
'*****************************************************************
'Puts a object in a User's slot from the current User's position
'*****************************************************************

Dim X As Integer
Dim Y As Integer
Dim Slot As Byte

X = UserList(UserIndex).Pos.X
Y = UserList(UserIndex).Pos.Y

'Check for object on ground
If MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.ObjIndex <= 0 Then
    Call SendData(ToIndex, UserIndex, 0, "@Nothing there." & FONTTYPE_INFO)
    Exit Sub
End If

'Check to see if User already has object type
Slot = 1
Do Until UserList(UserIndex).Object(Slot).ObjIndex = MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.ObjIndex
    Slot = Slot + 1

    If Slot > MAX_INVENTORY_SLOTS Then
        Exit Do
    End If
Loop

'Else check if there is a empty slot
If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1
        Do Until UserList(UserIndex).Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > MAX_INVENTORY_SLOTS Then
                Call SendData(ToIndex, UserIndex, 0, "@Can't Hold anymore." & FONTTYPE_INFO)
                Exit Sub
                Exit Do
            End If
        Loop
End If

'Fill object slot
If UserList(UserIndex).Object(Slot).Amount + MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.Amount <= MAX_INVENTORY_OBJS Then
    'Under MAX_INV_OBJS
    UserList(UserIndex).Object(Slot).ObjIndex = MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.ObjIndex
    UserList(UserIndex).Object(Slot).Amount = UserList(UserIndex).Object(Slot).Amount + MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.Amount
    Call EraseObj(ToMap, 0, UserList(UserIndex).Pos.Map, MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.Amount, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
Else
    'Over MAX_INV_OBJS
    If MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.Amount < UserList(UserIndex).Object(Slot).Amount Then
        MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.Amount = Abs(MAX_INVENTORY_OBJS - (UserList(UserIndex).Object(Slot).Amount + MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.Amount))
    Else
        MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.Amount = Abs((MAX_INVENTORY_OBJS + UserList(UserIndex).Object(Slot).Amount) - MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.Amount)
    End If
    UserList(UserIndex).Object(Slot).Amount = MAX_INVENTORY_OBJS
End If

Call UpdateUserInv(False, UserIndex, Slot)

End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
'*****************************************************************
'Updates a User's inventory
'*****************************************************************
Dim NullObj As UserOBJ
Dim LoopC As Byte

'Update one slot
If UpdateAll = False Then

    'Update User inventory
    If UserList(UserIndex).Object(Slot).ObjIndex > 0 Then
        Call ChangeUserInv(UserIndex, Slot, UserList(UserIndex).Object(Slot))
    Else
        Call ChangeUserInv(UserIndex, Slot, NullObj)
    End If

Else

'Update every slot
    For LoopC = 1 To MAX_INVENTORY_SLOTS

        'Update User invetory
        If UserList(UserIndex).Object(LoopC).ObjIndex > 0 Then
            Call ChangeUserInv(UserIndex, LoopC, UserList(UserIndex).Object(LoopC))
        Else
            Call ChangeUserInv(UserIndex, LoopC, NullObj)
        End If

    Next LoopC

End If

End Sub

Sub ChangeUserInv(ByVal UserIndex As Integer, ByVal Slot As Byte, Object As UserOBJ)
'*****************************************************************
'Changes a user's inventory
'*****************************************************************

UserList(UserIndex).Object(Slot) = Object

If Object.ObjIndex > 0 Then

    Call SendData(ToIndex, UserIndex, 0, "SIS" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GRHIndex)

Else

    Call SendData(ToIndex, UserIndex, 0, "SIS" & Slot & "," & "0" & "," & "(None)" & "," & "0" & "," & "0")

End If


End Sub

Sub DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'*****************************************************************
'Drops a object from a User's slot
'*****************************************************************
Dim Obj As Obj

'Check amount
If Num <= 0 Then
    Exit Sub
End If

If Num > UserList(UserIndex).Object(Slot).Amount Then
    Num = UserList(UserIndex).Object(Slot).Amount
End If

'Check for object on gorund
If MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.ObjIndex <> 0 Then
    Call SendData(ToIndex, UserIndex, 0, "@No room on ground." & FONTTYPE_INFO)
    Exit Sub
End If

Obj.ObjIndex = UserList(UserIndex).Object(Slot).ObjIndex
Obj.Amount = Num
Call MakeObj(ToMap, 0, Map, Obj, Map, X, Y)

'Remove object
UserList(UserIndex).Object(Slot).Amount = UserList(UserIndex).Object(Slot).Amount - Num
If UserList(UserIndex).Object(Slot).Amount <= 0 Then
    
    'Unequip is the object is currently equipped
    If UserList(UserIndex).Object(Slot).Equipped = 1 Then
        Call RemoveInvItem(UserIndex, Slot)
    End If
    
    UserList(UserIndex).Object(Slot).ObjIndex = 0
    UserList(UserIndex).Object(Slot).Amount = 0
    UserList(UserIndex).Object(Slot).Equipped = 0
End If

Call UpdateUserInv(False, UserIndex, Slot)

End Sub

Sub CloseNPC(ByVal NPCIndex As Integer)
'*****************************************************************
'Closes a NPC
'*****************************************************************

NPCList(NPCIndex).Flags.NPCActive = 0

'update last npc
If NPCIndex = LastNPC Then
    Do Until NPCList(LastNPC).Flags.NPCActive = 1
        LastNPC = LastNPC - 1
        If LastNPC = 0 Then Exit Do
    Loop
End If
  
'update number of users
If NumNPCs <> 0 Then
    NumNPCs = NumNPCs - 1
End If

End Sub

Sub UserAttackNPC(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)
'*****************************************************************
'Have a User attack a NPC
'*****************************************************************
Dim Hit As Integer

'Calculate hit
Hit = Int(RandomNumber(UserList(UserIndex).Stats.MinHIT, UserList(UserIndex).Stats.MaxHIT))
Hit = Hit - (NPCList(NPCIndex).Stats.DEF / 2)
If Hit < 1 Then Hit = 1

'Hit NPC
SendData ToIndex, UserIndex, 0, "@You hit the " & NPCList(NPCIndex).Name & " for " & Hit & "!" & FONTTYPE_FIGHT
NPCList(NPCIndex).Stats.MinHP = NPCList(NPCIndex).Stats.MinHP - Hit

'NPC Die
If NPCList(NPCIndex).Stats.MinHP <= 0 Then
    
    'Give EXP and gold
    UserList(UserIndex).Stats.EXP = UserList(UserIndex).Stats.EXP + NPCList(NPCIndex).GiveEXP
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + NPCList(NPCIndex).GiveGLD

    'Kill it
    SendData ToIndex, UserIndex, 0, "@You kill the " & NPCList(NPCIndex).Name & "!" & FONTTYPE_FIGHT
    KillNPC NPCIndex

End If

'Check user for level up
CheckUserLevel UserIndex
'Set update stats flag
UserList(UserIndex).Flags.StatsChanged = 1

End Sub

Sub NPCAttackUser(ByVal NPCIndex As Integer, ByVal UserIndex As Integer)
'*****************************************************************
'Have a NPC attack a User
'*****************************************************************
Dim Hit As Integer

'Don't allow if switchingmaps maps
If UserList(UserIndex).Flags.SwitchingMaps Then
    Exit Sub
End If

'Calculate hit
Hit = Int(RandomNumber(NPCList(NPCIndex).Stats.MinHIT, NPCList(NPCIndex).Stats.MaxHIT))
Hit = Hit - (UserList(UserIndex).Stats.DEF / 2)
If Hit < 1 Then Hit = 1

'Hit user
SendData ToIndex, UserIndex, 0, "@" & NPCList(NPCIndex).Name & " hits you for " & Hit & "!" & FONTTYPE_FIGHT
UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - Hit

'User Die
If UserList(UserIndex).Stats.MinHP <= 0 Then
    
    'Kill user
    SendData ToIndex, UserIndex, 0, "@The " & NPCList(NPCIndex).Name & " kills you!" & FONTTYPE_FIGHT
    KillUser UserIndex

End If

'Set update stats flag
UserList(UserIndex).Flags.StatsChanged = 1

End Sub

Sub UserAttackUser(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)
'*****************************************************************
'Have a user attack a user
'*****************************************************************
Dim Hit As Integer

'Don't allow if switchingmaps maps
If UserList(VictimIndex).Flags.SwitchingMaps Then
    Exit Sub
End If

'Calculate hit
Hit = Int(RandomNumber(UserList(AttackerIndex).Stats.MinHIT, UserList(AttackerIndex).Stats.MaxHIT))
Hit = Hit - (UserList(VictimIndex).Stats.DEF / 2)
If Hit < 1 Then Hit = 1

'Hit User
SendData ToIndex, AttackerIndex, 0, "@You hit " & UserList(VictimIndex).Name & " for " & Hit & "!" & FONTTYPE_FIGHT
SendData ToIndex, VictimIndex, 0, "@" & UserList(AttackerIndex).Name & " hits you for " & Hit & "!" & FONTTYPE_FIGHT
UserList(VictimIndex).Stats.MinHP = UserList(VictimIndex).Stats.MinHP - Hit

'User Die
If UserList(VictimIndex).Stats.MinHP <= 0 Then
    
    'Give EXP and gold
    UserList(AttackerIndex).Stats.EXP = UserList(AttackerIndex).Stats.EXP + (UserList(VictimIndex).Stats.ELV * 20)

    'Kill user
    SendData ToIndex, AttackerIndex, 0, "@You kill " & UserList(VictimIndex).Name & "!" & FONTTYPE_FIGHT
    SendData ToIndex, VictimIndex, 0, "@" & UserList(AttackerIndex).Name & " kills you!" & FONTTYPE_FIGHT
    KillUser VictimIndex

End If

'update users level and stats

CheckUserLevel AttackerIndex
'Set update stats flag
UserList(AttackerIndex).Flags.StatsChanged = 1

CheckUserLevel VictimIndex
'Set update stats flag
UserList(VictimIndex).Flags.StatsChanged = 1

End Sub

Sub UserAttack(ByVal UserIndex As Integer)
'*****************************************************************
'Begin a user attack sequence
'*****************************************************************
Dim AttackPos As WorldPos

'Check switching maps
If UserList(UserIndex).Flags.SwitchingMaps Then
    Exit Sub
End If

'Check attacker counter
If UserList(UserIndex).Counters.AttackCounter > 0 Then
    Exit Sub
End If

'Check stanima
If UserList(UserIndex).Stats.MinSTA <= 0 Then
    Exit Sub
End If

'update counters
UserList(UserIndex).Counters.AttackCounter = STAT_ATTACKWAIT
UserList(UserIndex).Stats.MinSTA = UserList(UserIndex).Stats.MinSTA - 1

'Get tile user is attacking
AttackPos = UserList(UserIndex).Pos
HeadtoPos UserList(UserIndex).Char.Heading, AttackPos

'Play attack sound
SendData ToPCArea, UserIndex, AttackPos.Map, "PLW" & SOUND_SWING

'Exit if not legal
If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
    Exit Sub
End If

'Look for user
If MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex > 0 Then
    UserAttackUser UserIndex, MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex
    Exit Sub
End If

'Look for NPC
If MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NPCIndex > 0 Then

    If NPCList(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NPCIndex).Attackable Then
        UserAttackNPC UserIndex, MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NPCIndex
    Else
        SendData ToIndex, UserIndex, 0, "@A mysterious force prevents you from attacking..." & FONTTYPE_FIGHT
    End If

    Exit Sub
End If

End Sub

Function UserIndex(ByVal SocketId As Integer) As Integer
'*****************************************************************
'Finds the User with a certain SocketID
'*****************************************************************
Dim LoopC As Integer
  
LoopC = 1
  
Do Until UserList(LoopC).ConnID = SocketId

    LoopC = LoopC + 1
    
    If LoopC > MaxUsers Then
        UserIndex = 0
        Exit Function
    End If
    
Loop
  
UserIndex = LoopC

End Function

Function CheckForSameIP(ByVal UserIndex As Integer, ByVal UserIP As String) As Boolean
'*****************************************************************
'Checks for a user with the same IP
'*****************************************************************
Dim LoopC As Integer

For LoopC = 1 To LastUser

    If UserList(LoopC).Flags.UserLogged = 1 Then
        If UserList(LoopC).IP = UserIP And UserIndex <> LoopC Then
            CheckForSameIP = True
            Exit Function
        End If
    End If

Next LoopC

CheckForSameIP = False

End Function

Function CheckForSameName(ByVal UserIndex As Integer, ByVal Name As String) As Boolean
'*****************************************************************
'Checks for a user with the same Name
'*****************************************************************
Dim LoopC As Integer

For LoopC = 1 To LastUser

    If UserList(LoopC).Flags.UserLogged = 1 Then
        If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserIndex <> LoopC Then
            CheckForSameName = True
            Exit Function
        End If
    End If

Next LoopC

CheckForSameName = False

End Function

Sub HeadtoPos(ByVal Head As Byte, ByRef Pos As WorldPos)
'*****************************************************************
'Takes Pos and ad moves it in heading direction
'*****************************************************************
Dim X As Integer
Dim Y As Integer
Dim tempVar As Single
Dim nX As Integer
Dim nY As Integer

X = Pos.X
Y = Pos.Y

If Head = NORTH Then
    nX = X
    nY = Y - 1
End If

If Head = SOUTH Then
    nX = X
    nY = Y + 1
End If

If Head = EAST Then
    nX = X + 1
    nY = Y
End If

If Head = WEST Then
    nX = X - 1
    nY = Y
End If

'return values
Pos.X = nX
Pos.Y = nY

End Sub

Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
'*****************************************************************
'Sends a user's stats to text window
'*****************************************************************

Call SendData(ToIndex, sendIndex, 0, "@Stats for: " & UserList(UserIndex).Name & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "@Level: " & UserList(UserIndex).Stats.ELV & "  EXP: " & UserList(UserIndex).Stats.EXP & "/" & UserList(UserIndex).Stats.ELU & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "@Fitness: " & UserList(UserIndex).Stats.FIT & "  Metabolism: " & UserList(UserIndex).Stats.MET & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "@Health: " & UserList(UserIndex).Stats.MinHP & "/" & UserList(UserIndex).Stats.MaxHP & "  Mana: " & UserList(UserIndex).Stats.MinMAN & "/" & UserList(UserIndex).Stats.MaxMAN & "  Stamina: " & UserList(UserIndex).Stats.MinSTA & "/" & UserList(UserIndex).Stats.MaxSTA & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "@Min Hit/Max Hit: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHIT & "   Defense: " & UserList(UserIndex).Stats.DEF & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "@Gold: " & UserList(UserIndex).Stats.GLD & "  Position: " & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y & " in map " & UserList(UserIndex).Pos.Map & FONTTYPE_INFO)

End Sub

Sub UpdateUserMap(ByVal UserIndex As Integer)
'*****************************************************************
'Updates a user with the place of all chars in the Map
'*****************************************************************
Dim Map As Integer
Dim X As Integer
Dim Y As Integer

Map = UserList(UserIndex).Pos.Map

'Place chars
For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        If MapData(Map, X, Y).UserIndex > 0 Then
            Call MakeUserChar(ToIndex, UserIndex, 0, MapData(Map, X, Y).UserIndex, Map, X, Y)
        End If

        If MapData(Map, X, Y).NPCIndex > 0 Then
            Call MakeNPCChar(ToIndex, UserIndex, 0, MapData(Map, X, Y).NPCIndex, Map, X, Y)
        End If

        If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
            Call MakeObj(ToIndex, UserIndex, 0, MapData(Map, X, Y).ObjInfo, Map, X, Y)
        End If

    Next X
Next Y

End Sub

Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As Byte)
'*****************************************************************
'Moves a User from one tile to another
'*****************************************************************
Dim nPos As WorldPos

'Move
nPos = UserList(UserIndex).Pos
Call HeadtoPos(nHeading, nPos)

'Move if legal pos
If LegalPos(UserList(UserIndex).Pos.Map, nPos.X, nPos.Y) = True Then
    Call SendData(ToMapButIndex, UserIndex, UserList(UserIndex).Pos.Map, "MOC" & UserList(UserIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y)

    'Update map and user pos
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
    UserList(UserIndex).Pos = nPos
    UserList(UserIndex).Char.Heading = nHeading
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = UserIndex
Else
    'else correct user's pos
    Call SendData(ToIndex, UserIndex, 0, "SUP" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
End If

End Sub

Sub MoveNPCChar(ByVal NPCIndex As Integer, ByVal nHeading As Byte)
'*****************************************************************
'Moves a NPC from one tile to another
'*****************************************************************
Dim nPos As WorldPos

'Move
nPos = NPCList(NPCIndex).Pos
Call HeadtoPos(nHeading, nPos)

'Move if legal pos
If LegalPos(NPCList(NPCIndex).Pos.Map, nPos.X, nPos.Y) = True Then
    Call SendData(ToMap, 0, NPCList(NPCIndex).Pos.Map, "MOC" & NPCList(NPCIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y)

    'Update map and user pos
    MapData(NPCList(NPCIndex).Pos.Map, NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y).NPCIndex = 0
    NPCList(NPCIndex).Pos = nPos
    NPCList(NPCIndex).Char.Heading = nHeading
    MapData(NPCList(NPCIndex).Pos.Map, NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y).NPCIndex = NPCIndex
End If

End Sub

Sub MakeUserChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'*****************************************************************
'Makes and places a user's character
'*****************************************************************
Dim CharIndex As Integer

'If needed make a new character in list
If UserList(UserIndex).Char.CharIndex = 0 Then
    CharIndex = NextOpenCharIndex
    UserList(UserIndex).Char.CharIndex = CharIndex
    CharList(CharIndex) = UserIndex
End If

'Place character on map
MapData(Map, X, Y).UserIndex = UserIndex

'Send make character command to clients
Call SendData(sndRoute, sndIndex, sndMap, "MAC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y)

End Sub

Sub MakeNPCChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal NPCIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'*****************************************************************
'Makes and places a NPC character
'*****************************************************************
Dim CharIndex As Integer

'If needed make a new character in list
If NPCList(NPCIndex).Char.CharIndex = 0 Then
    CharIndex = NextOpenCharIndex
    NPCList(NPCIndex).Char.CharIndex = CharIndex
    CharList(CharIndex) = NPCIndex
End If

'Place character on map
MapData(Map, X, Y).NPCIndex = NPCIndex

'Set alive flag
NPCList(NPCIndex).Flags.NPCAlive = 1

'Send make character command to clients
Call SendData(sndRoute, sndIndex, sndMap, "MAC" & NPCList(NPCIndex).Char.Body & "," & NPCList(NPCIndex).Char.Head & "," & NPCList(NPCIndex).Char.Heading & "," & NPCList(NPCIndex).Char.CharIndex & "," & X & "," & Y)

End Sub

Function LegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************

'Make sure it's a legal map
If Map <= 0 Or Map > NumMaps Then
    LegalPos = False
    Exit Function
End If

'Check to see if its out of bounds
If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    LegalPos = False
    Exit Function
End If

'Check to see if its blocked
If MapData(Map, X, Y).Blocked = 1 Then
    LegalPos = False
    Exit Function
End If

'User
If MapData(Map, X, Y).UserIndex > 0 Then
    LegalPos = False
    Exit Function
End If

'NPC
If MapData(Map, X, Y).NPCIndex > 0 Then
    LegalPos = False
    Exit Function
End If

LegalPos = True

End Function

Sub SendHelp(ByVal UserIndex As Integer)
'*****************************************************************
'Sends help strings to Index
'*****************************************************************
Dim NumHelpLines As Integer
Dim LoopC As Integer

NumHelpLines = Val(GetVar(IniPath & "Help.dat", "INIT", "NumLines"))

For LoopC = 1 To NumHelpLines
    Call SendData(ToIndex, UserIndex, 0, "@" & GetVar(IniPath & "Help.dat", "Help", "Line" & LoopC) & FONTTYPE_INFO)
Next LoopC

End Sub

Sub EraseUserChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal UserIndex As Integer)
'*****************************************************************
'Erase a character
'*****************************************************************

'Remove from list
CharList(UserList(UserIndex).Char.CharIndex) = 0

'Update LsstChar
If UserList(UserIndex).Char.CharIndex = LastChar Then
    Do Until CharList(LastChar) > 0
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If

'Remove from map
MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0

'Send erase command to clients
Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "ERC" & UserList(UserIndex).Char.CharIndex)

'Update userlist
UserList(UserIndex).Char.CharIndex = 0

'update NumChars
NumChars = NumChars - 1

End Sub

Sub EraseNPCChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal NPCIndex As Integer)
'*****************************************************************
'Erase a character
'*****************************************************************

'Remove from list
CharList(NPCList(NPCIndex).Char.CharIndex) = 0

'Update LastChar
If NPCList(NPCIndex).Char.CharIndex = LastChar Then
    Do Until CharList(LastChar) > 0
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If

'Remove from map
MapData(NPCList(NPCIndex).Pos.Map, NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y).NPCIndex = 0

'Send erase command to clients
Call SendData(ToMap, 0, NPCList(NPCIndex).Pos.Map, "ERC" & NPCList(NPCIndex).Char.CharIndex)

'Update npclist
NPCList(NPCIndex).Char.CharIndex = 0

'Set alive flag
NPCList(NPCIndex).Flags.NPCAlive = 0

'update NumChars
NumChars = NumChars - 1

End Sub

Sub LookatTile(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'*****************************************************************
'Responds to the user clicking on a square
'*****************************************************************
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer

'Check if legal
If InMapBounds(Map, X, Y) = False Then
    Exit Sub
End If

'*** Check for object ***
If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
    Call SendData(ToIndex, UserIndex, 0, "@You see a " & ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Name & FONTTYPE_TALK)
    FoundSomething = 1
End If

'*** Check for Characters ***
If Y + 1 <= YMaxMapSize Then
    If MapData(Map, X, Y + 1).UserIndex > 0 Then
        TempCharIndex = MapData(Map, X, Y + 1).UserIndex
        FoundChar = 1
    End If
    If MapData(Map, X, Y + 1).NPCIndex > 0 Then
        TempCharIndex = MapData(Map, X, Y + 1).NPCIndex
        FoundChar = 2
    End If
End If
'Check for Character
If FoundChar = 0 Then
    If MapData(Map, X, Y).UserIndex > 0 Then
        TempCharIndex = MapData(Map, X, Y).UserIndex
        FoundChar = 1
    End If
    If MapData(Map, X, Y).NPCIndex > 0 Then
        TempCharIndex = MapData(Map, X, Y).NPCIndex
        FoundChar = 2
    End If
End If
'React to character
If FoundChar = 1 Then
        If Len(UserList(TempCharIndex).Desc) > 1 Then
            Call SendData(ToIndex, UserIndex, 0, "@You see " & UserList(TempCharIndex).modName & " - " & UserList(TempCharIndex).Desc)
        Else
            Call SendData(ToIndex, UserIndex, 0, "@You see " & UserList(TempCharIndex).modName & ".")
        End If
        FoundSomething = 1
End If
If FoundChar = 2 Then
        If Len(UserList(TempCharIndex).Desc) > 1 Then
            Call SendData(ToIndex, UserIndex, 0, "@You see " & NPCList(TempCharIndex).Name & " - " & NPCList(TempCharIndex).Desc)
        Else
            Call SendData(ToIndex, UserIndex, 0, "@You see " & NPCList(TempCharIndex).Name & ".")
        End If
        FoundSomething = 1
End If

'*** Didn't find anything ***
If FoundSomething = 0 Then
    Call SendData(ToIndex, UserIndex, 0, "@You see nothing of interest.")
End If

End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'*****************************************************************
'Warps user to another spot
'*****************************************************************
Dim OldMap As Integer
Dim OldX As Integer
Dim OldY As Integer

OldMap = UserList(UserIndex).Pos.Map
OldX = UserList(UserIndex).Pos.X
OldY = UserList(UserIndex).Pos.Y

Call EraseUserChar(ToMap, 0, OldMap, UserIndex)

UserList(UserIndex).Pos.X = X
UserList(UserIndex).Pos.Y = Y
UserList(UserIndex).Pos.Map = Map

If OldMap <> Map Then
    'Set switchingmap flag
    UserList(UserIndex).Flags.SwitchingMaps = 1
    
    'Tell client to try switching maps
    Call SendData(ToIndex, UserIndex, 0, "SCM" & Map & "," & MapInfo(Map).MapVersion)

    'Update new Map Users
    MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
    'Update old Map Users
    MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
    If MapInfo(OldMap).NumUsers < 0 Then
        MapInfo(OldMap).NumUsers = 0
    End If
    
    'Show Character to others
    Call MakeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    
Else
    
    Call MakeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    Call SendData(ToIndex, UserIndex, 0, "SUC" & UserList(UserIndex).Char.CharIndex)

End If

End Sub

Sub SendUserStatsBox(ByVal UserIndex As Integer)
'*****************************************************************
'Updates a User's stat box
'*****************************************************************

Call SendData(ToIndex, UserIndex, 0, "SST" & UserList(UserIndex).Stats.MaxHP & "," & UserList(UserIndex).Stats.MinHP & "," & UserList(UserIndex).Stats.MaxMAN & "," & UserList(UserIndex).Stats.MinMAN & "," & UserList(UserIndex).Stats.MaxSTA & "," & UserList(UserIndex).Stats.MinSTA & "," & UserList(UserIndex).Stats.GLD & "," & UserList(UserIndex).Stats.ELV)

End Sub

Function FindDirection(Pos As WorldPos, Target As WorldPos) As Byte
'*****************************************************************
'Returns the direction in which the Target is from the Pos, 0 if equal
'*****************************************************************
Dim X As Integer
Dim Y As Integer

X = Pos.X - Target.X
Y = Pos.Y - Target.Y

'NE
If Sgn(X) = -1 And Sgn(Y) = 1 Then
    FindDirection = NORTH
    Exit Function
End If

'NW
If Sgn(X) = 1 And Sgn(Y) = 1 Then
    FindDirection = WEST
    Exit Function
End If

'SW
If Sgn(X) = 1 And Sgn(Y) = -1 Then
    FindDirection = WEST
    Exit Function
End If

'SE
If Sgn(X) = -1 And Sgn(Y) = -1 Then
    FindDirection = SOUTH
    Exit Function
End If

'South
If Sgn(X) = 0 And Sgn(Y) = -1 Then
    FindDirection = SOUTH
    Exit Function
End If

'north
If Sgn(X) = 0 And Sgn(Y) = 1 Then
    FindDirection = NORTH
    Exit Function
End If

'West
If Sgn(X) = 1 And Sgn(Y) = 0 Then
    FindDirection = WEST
    Exit Function
End If

'East
If Sgn(X) = -1 And Sgn(Y) = 0 Then
    FindDirection = EAST
    Exit Function
End If

'Same spot
If Sgn(X) = 0 And Sgn(Y) = 0 Then
    FindDirection = 0
    Exit Function
End If

End Function



