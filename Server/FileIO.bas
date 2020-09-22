Attribute VB_Name = "FileIO"
Option Explicit

Sub LoadOBJData()
'*****************************************************************
'Setup OBJ list
'*****************************************************************
Dim Object As Integer

'Get Number of Objects
NumObjDatas = Val(GetVar(IniPath & "Obj.dat", "INIT", "NumObjs"))
ReDim ObjData(1 To NumObjDatas) As ObjData
  
'Fill Object List
For Object = 1 To NumObjDatas
    
    ObjData(Object).Name = GetVar(IniPath & "Obj.dat", "OBJ" & Object, "Name")
    
    ObjData(Object).GRHIndex = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "GrhIndex"))
    
    ObjData(Object).ObjType = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "ObjType"))

    
    ObjData(Object).MaxHIT = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "MaxHIT"))
    ObjData(Object).MinHIT = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "MinHIT"))
    ObjData(Object).MaxHP = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "MaxHP"))
    ObjData(Object).MinHP = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "MinHP"))
 
    ObjData(Object).DEF = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "DEF"))

Next Object

End Sub

Sub LoadUserStats(UserIndex As Integer, UserFile As String)
'*****************************************************************
'Loads a user's stats from a text file
'*****************************************************************

UserList(UserIndex).Stats.GLD = Val(GetVar(UserFile, "STATS", "GLD"))

UserList(UserIndex).Stats.MET = Val(GetVar(UserFile, "STATS", "MET"))
UserList(UserIndex).Stats.MaxHP = Val(GetVar(UserFile, "STATS", "MaxHP"))
UserList(UserIndex).Stats.MinHP = Val(GetVar(UserFile, "STATS", "MinHP"))

UserList(UserIndex).Stats.FIT = Val(GetVar(UserFile, "STATS", "FIT"))
UserList(UserIndex).Stats.MinSTA = Val(GetVar(UserFile, "STATS", "MinSTA"))
UserList(UserIndex).Stats.MaxSTA = Val(GetVar(UserFile, "STATS", "MaxSTA"))

UserList(UserIndex).Stats.MaxMAN = Val(GetVar(UserFile, "STATS", "MaxMAN"))
UserList(UserIndex).Stats.MinMAN = Val(GetVar(UserFile, "STATS", "MinMAN"))

UserList(UserIndex).Stats.MaxHIT = Val(GetVar(UserFile, "STATS", "MaxHIT"))
UserList(UserIndex).Stats.MinHIT = Val(GetVar(UserFile, "STATS", "MinHIT"))
UserList(UserIndex).Stats.DEF = Val(GetVar(UserFile, "STATS", "DEF"))

UserList(UserIndex).Stats.EXP = Val(GetVar(UserFile, "STATS", "EXP"))
UserList(UserIndex).Stats.ELU = Val(GetVar(UserFile, "STATS", "ELU"))
UserList(UserIndex).Stats.ELV = Val(GetVar(UserFile, "STATS", "ELV"))

End Sub

Sub LoadUserInit(UserIndex As Integer, UserFile As String)
'*****************************************************************
'Loads the user's Init stuff
'*****************************************************************

Dim LoopC As Integer
Dim ln As String

'Get INIT
UserList(UserIndex).Char.Heading = Val(GetVar(UserFile, "INIT", "Heading"))
UserList(UserIndex).Char.Head = Val(GetVar(UserFile, "INIT", "Head"))
UserList(UserIndex).Char.Body = Val(GetVar(UserFile, "INIT", "Body"))
UserList(UserIndex).Desc = GetVar(UserFile, "INIT", "Desc")

'Get last postion
UserList(UserIndex).Pos.Map = Val(ReadField(1, GetVar(UserFile, "INIT", "Position"), 45))
UserList(UserIndex).Pos.X = Val(ReadField(2, GetVar(UserFile, "INIT", "Position"), 45))
UserList(UserIndex).Pos.Y = Val(ReadField(3, GetVar(UserFile, "INIT", "Position"), 45))

'Get object list
For LoopC = 1 To MAX_INVENTORY_SLOTS
    ln = GetVar(UserFile, "Inventory", "Obj" & LoopC)
    UserList(UserIndex).Object(LoopC).ObjIndex = Val(ReadField(1, ln, 45))
    UserList(UserIndex).Object(LoopC).Amount = Val(ReadField(2, ln, 45))
    UserList(UserIndex).Object(LoopC).Equipped = Val(ReadField(3, ln, 45))
Next LoopC

'Get Weapon objectindex and slot
UserList(UserIndex).WeaponEqpSlot = Val(GetVar(UserFile, "Inventory", "WeaponEqpSlot"))
If UserList(UserIndex).WeaponEqpSlot > 0 Then
    UserList(UserIndex).WeaponEqpObjIndex = UserList(UserIndex).Object(UserList(UserIndex).WeaponEqpSlot).ObjIndex
End If

'Get Armour objectindex and slot
UserList(UserIndex).ArmourEqpSlot = Val(GetVar(UserFile, "Inventory", "ArmourEqpSlot"))
If UserList(UserIndex).ArmourEqpSlot > 0 Then
    UserList(UserIndex).ArmourEqpObjIndex = UserList(UserIndex).Object(UserList(UserIndex).ArmourEqpSlot).ObjIndex
End If

End Sub

Function WizCheck(Name As String) As Boolean
'*****************************************************************
'Checks to see if Name is a wizard
'*****************************************************************
Dim NumWizs As Integer
Dim WizNum As Integer

NumWizs = Val(GetVar(IniPath & "Server.ini", "INIT", "NumWizs"))
For WizNum = 1 To NumWizs
    If UCase(Name) = UCase(GetVar(IniPath & "Server.ini", "WizList", "wiz" & WizNum)) Then
        WizCheck = True
        Exit Function
    End If
Next WizNum

WizCheck = False

End Function

Function GetVar(File As String, Main As String, Var As String) As String
'*****************************************************************
'Gets a variable from a text file
'*****************************************************************
Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
  
szReturn = ""
  
sSpaces = Space(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish
  
  
getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File
  
GetVar = RTrim(sSpaces)
GetVar = Left(GetVar, Len(GetVar) - 1)
  
End Function

Sub LoadMapData()
'*****************************************************************
'Loads the MapX.X files
'*****************************************************************
Dim Map As Integer
Dim LoopC As Integer
Dim X As Integer
Dim Y As Integer
Dim TempInt As Integer

NumMaps = Val(GetVar(IniPath & "Map.dat", "INIT", "NumMaps"))
MapPath = GetVar(IniPath & "Map.dat", "INIT", "MapPath")

ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
ReDim MapInfo(1 To NumMaps) As MapInfo
  
For Map = 1 To NumMaps
   
    'Open files
    
    'map
    Open App.Path & MapPath & "Map" & Map & ".map" For Binary As #1
    Seek #1, 1
    
    'inf
    Open App.Path & MapPath & "Map" & Map & ".inf" For Binary As #2
    Seek #2, 1
    
    'map Header
    Get #1, , MapInfo(Map).MapVersion
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

            '.dat file
            Get #1, , MapData(Map, X, Y).Blocked
            
            'Get GRH number
            For LoopC = 1 To 4
                Get #1, , MapData(Map, X, Y).Graphic(LoopC)
            Next LoopC
            
            'Space holder for future expansion
            Get #1, , TempInt
            Get #1, , TempInt
                                
                                
            '.inf file
            
            'tile exit
            Get #2, , MapData(Map, X, Y).TileExit.Map
            Get #2, , MapData(Map, X, Y).TileExit.X
            Get #2, , MapData(Map, X, Y).TileExit.Y
            
            'Get and make NPC
            Get #2, , TempInt
            If TempInt > 0 Then
                SpawnNPC OpenNPC(TempInt), Map, X, Y
            Else
                MapData(Map, X, Y).NPCIndex = 0
            End If

            'Get and make Object
            Get #2, , MapData(Map, X, Y).ObjInfo.ObjIndex
            Get #2, , MapData(Map, X, Y).ObjInfo.Amount

            'Space holder for future expansion
            Get #2, , TempInt
            Get #2, , TempInt
        
        Next X
    Next Y

    'Close files
    Close #1
    Close #2

    'Other Room Data
    MapInfo(Map).Name = GetVar(App.Path & MapPath & "Map" & Map & ".dat", "Map" & Map, "Name")
    MapInfo(Map).Music = GetVar(App.Path & MapPath & "Map" & Map & ".dat", "Map" & Map, "MusicNum")
    MapInfo(Map).StartPos.Map = Val(ReadField(1, GetVar(App.Path & MapPath & "Map" & Map & ".dat", "Map" & Map, "StartPos"), 45))
    MapInfo(Map).StartPos.X = Val(ReadField(2, GetVar(App.Path & MapPath & "Map" & Map & ".dat", "Map" & Map, "StartPos"), 45))
    MapInfo(Map).StartPos.Y = Val(ReadField(3, GetVar(App.Path & MapPath & "Map" & Map & ".dat", "Map" & Map, "StartPos"), 45))

Next Map

End Sub

Sub LoadSini()
'*****************************************************************
'Loads the Server.ini
'*****************************************************************

'Misc
frmMain.txPortNumber.Text = GetVar(IniPath & "Server.ini", "INIT", "StartPort")
HideMe = Val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))
AllowMultiLogins = Val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
IdleLimit = Val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))

'Start pos
StartPos.Map = Val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))
StartPos.X = Val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))
StartPos.Y = Val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))

'Res pos
ResPos.Map = Val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.X = Val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.Y = Val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))

'Ressurect pos
ResPos.Map = Val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.X = Val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.Y = Val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
  
'Max users
MaxUsers = Val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))
ReDim UserList(1 To MaxUsers) As User

End Sub

Sub WriteVar(File As String, Main As String, Var As String, Value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************

writeprivateprofilestring Main, Var, Value, File
    
End Sub

Sub SaveUser(UserIndex As Integer, UserFile As String)
'*****************************************************************
'Saves a user's data to a .chr file
'*****************************************************************
Dim LoopC As Integer

Call WriteVar(UserFile, "INIT", "Password", UserList(UserIndex).Password)
Call WriteVar(UserFile, "INIT", "Desc", UserList(UserIndex).Desc)
Call WriteVar(UserFile, "INIT", "Heading", Str(UserList(UserIndex).Char.Heading))
Call WriteVar(UserFile, "INIT", "Head", Str(UserList(UserIndex).Char.Head))
Call WriteVar(UserFile, "INIT", "Body", Str(UserList(UserIndex).Char.Body))

Call WriteVar(UserFile, "INIT", "LastIP", UserList(UserIndex).IP)
Call WriteVar(UserFile, "INIT", "Position", UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.X & "-" & UserList(UserIndex).Pos.Y)


Call WriteVar(UserFile, "STATS", "GLD", Str(UserList(UserIndex).Stats.GLD))

Call WriteVar(UserFile, "STATS", "MET", Str(UserList(UserIndex).Stats.MET))
Call WriteVar(UserFile, "STATS", "MaxHP", Str(UserList(UserIndex).Stats.MaxHP))
Call WriteVar(UserFile, "STATS", "MinHP", Str(UserList(UserIndex).Stats.MinHP))

Call WriteVar(UserFile, "STATS", "FIT", Str(UserList(UserIndex).Stats.FIT))
Call WriteVar(UserFile, "STATS", "MaxSTA", Str(UserList(UserIndex).Stats.MaxSTA))
Call WriteVar(UserFile, "STATS", "MinSTA", Str(UserList(UserIndex).Stats.MinSTA))

Call WriteVar(UserFile, "STATS", "MaxMAN", Str(UserList(UserIndex).Stats.MaxMAN))
Call WriteVar(UserFile, "STATS", "MinMAN", Str(UserList(UserIndex).Stats.MinMAN))

Call WriteVar(UserFile, "STATS", "MaxHIT", Str(UserList(UserIndex).Stats.MaxHIT))
Call WriteVar(UserFile, "STATS", "MinHIT", Str(UserList(UserIndex).Stats.MinHIT))
Call WriteVar(UserFile, "STATS", "DEF", Str(UserList(UserIndex).Stats.DEF))
  
Call WriteVar(UserFile, "STATS", "EXP", Str(UserList(UserIndex).Stats.EXP))
Call WriteVar(UserFile, "STATS", "ELV", Str(UserList(UserIndex).Stats.ELV))
Call WriteVar(UserFile, "STATS", "ELU", Str(UserList(UserIndex).Stats.ELU))
  
'Save Inv
For LoopC = 1 To MAX_INVENTORY_SLOTS
    Call WriteVar(UserFile, "Inventory", "Obj" & LoopC, UserList(UserIndex).Object(LoopC).ObjIndex & "-" & UserList(UserIndex).Object(LoopC).Amount & "-" & UserList(UserIndex).Object(LoopC).Equipped)
Next

'Write Weapon and Armour slots
Call WriteVar(UserFile, "Inventory", "WeaponEqpSlot", Str(UserList(UserIndex).WeaponEqpSlot))
Call WriteVar(UserFile, "Inventory", "ArmourEqpSlot", Str(UserList(UserIndex).ArmourEqpSlot))

End Sub



