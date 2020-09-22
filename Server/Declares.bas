Attribute VB_Name = "Declares"
Option Explicit

'********** Public CONSTANTS ***********

'Constants for Headings
Public Const NORTH = 1
Public Const EAST = 2
Public Const SOUTH = 3
Public Const WEST = 4

'Map sizes
Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

'Tile size in pixels
Public Const TileSizeX = 32
Public Const TileSizeY = 32

'Window size in tiles
Public Const XWindow = 17
Public Const YWindow = 13

'Sound constants
Public Const SOUND_BUMP = 1
Public Const SOUND_SWING = 2
Public Const SOUND_WARP = 3

'Object constants
Public Const MAX_INVENTORY_OBJS = 99
Public Const MAX_INVENTORY_SLOTS = 20

Public Const OBJTYPE_USEONCE = 1
Public Const OBJTYPE_WEAPON = 2
Public Const OBJTYPE_ARMOUR = 3

'Text type constants
Public Const FONTTYPE_TALK = "~0~0~0~0~0"
Public Const FONTTYPE_FIGHT = "~100~0~0~1~0"
Public Const FONTTYPE_WARNING = "~100~0~0~1~1"
Public Const FONTTYPE_INFO = "~0~100~0~0~0"

'Stat constants
Public Const STAT_MAXELV = 99
Public Const STAT_MAXHP = 999
Public Const STAT_MAXSTA = 999
Public Const STAT_MAXMAN = 999
Public Const STAT_MAXHIT = 99
Public Const STAT_MAXDEF = 99
Public Const STAT_MAXSTAT = 99     'Max for general stats (MET,FIT, ect)
Public Const STAT_METRATE = 50     'How many server ticks to recover some HP
Public Const STAT_FITRATE = 20     'How many server ticks to recover some STA
Public Const STAT_ATTACKWAIT = 5   'How many server ticks a user has to wait till he can attack again

'Other constants
Public Const MAX_CHARACTERS = 10000 'Should be max number users + max NPCs + some head room
Public Const MAX_NPCs = 5000 'How many NPCs are allowed in the game all together

'********** Public TYPES ***********

Type Position
    X As Integer
    Y As Integer
End Type

Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

'Holds data for a user or NPC character
Type Char
    CharIndex As Integer
    Head As Integer
    Body As Integer
    Heading As Byte
End Type

'** Object types **
Public Type ObjData
    Name As String
    ObjType As Integer
    GRHIndex As Integer
    MinHP As Integer
    MaxHP As Integer
    MinHIT As Integer
    MaxHIT As Integer
    DEF As Integer
End Type

Public Type Obj
    ObjIndex As Integer
    Amount As Integer
End Type

'** User Types **
'Stats for a user
Type UserStats
    GLD As Long
    MET As Integer
    MaxHP As Integer
    MinHP As Integer
    FIT As Integer
    MaxSTA As Integer
    MinSTA As Integer
    MaxMAN As Integer
    MinMAN As Integer
    MaxHIT As Integer
    MinHIT As Integer
    DEF As Integer
    EXP As Long
    ELV As Long
    ELU As Long
End Type

'Flags for a user
Type UserFlags
    UserLogged As Byte 'is the user logged in
    SwitchingMaps As Byte
    DownloadingMap As Byte
    ReadyForNextTile As Byte
    StatsChanged As Byte
End Type

Type UserCounters
    IdleCount As Long
    AttackCounter As Integer
    HPCounter As Integer
    STACounter As Integer
    SendMapCounter As WorldPos
End Type

Type UserOBJ
    ObjIndex As Integer
    Amount As Integer
    Equipped As Byte
End Type

'Holds data for a user
Type User
    Name As String
    modName As String
    Password As String
    Char As Char 'Defines users looks
    Desc As String
    
    Pos As WorldPos 'Current User Postion
    
    IP As String 'User Ip
    ConnID As Integer 'Connection ID
    RDBuffer As String 'Broken Line Buffer

    Object(1 To MAX_INVENTORY_SLOTS) As UserOBJ
    WeaponEqpObjIndex As Integer
    WeaponEqpSlot As Byte
    ArmourEqpObjIndex As Integer
    ArmourEqpSlot As Byte

    Counters As UserCounters
    Stats As UserStats
    Flags As UserFlags
End Type

'** NPC Types **
Type NPCStats
    MaxHP As Integer
    MinHP As Integer
    MaxHIT As Integer
    MinHIT As Integer
    DEF As Integer
End Type

Type NPCFlags
    NPCAlive As Byte  'is the NPC visible (plotted on map)
    NPCActive As Byte 'is the NPC being updated
End Type

Type NPCCounters
    RespawnCounter As Integer
End Type

Type NPC
    Name As String
    Char As Char 'Defines NPC looks
    Desc As String
    
    Pos As WorldPos 'Current NPC Postion
    StartPos As WorldPos
    
    Movement As Integer
    RespawnWait As Integer
    Attackable As Byte
    Hostile As Byte
    
    GiveEXP As Integer
    GiveGLD As Long
    
    Stats As NPCStats
    Flags As NPCFlags
    Counters As NPCCounters
End Type

'** Map Types **
'Tile Data
Type MapBlock
    Blocked As Byte
    Graphic(1 To 4) As Integer
    UserIndex As Integer
    NPCIndex As Integer
    ObjInfo As Obj
    TileExit As WorldPos
End Type

'Map info
Type MapInfo
    NumUsers As Integer
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
End Type

'********** Public VARS ***********

Public ENDL As String
Public ENDC As String

'Paths
Public IniPath As String
Public CharPath As String
Public MapPath As String

'Where the map borders are.. Set during load
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

Public ResPos As WorldPos 'Ressurect pos
Public StartPos As WorldPos 'Starting Pos (Loaded from Server.ini)


Public NumUsers As Integer 'current Number of Users
Public LastUser As Integer 'current Last User index
Public LastChar As Integer
Public NumChars As Integer
Public LastNPC As Integer
Public NumNPCs As Integer
Public NumMaps As Integer
Public NumObjDatas As Integer

Public AllowMultiLogins As Byte
Public IdleLimit As Integer
Public MaxUsers As Integer
Public HideMe As Byte

'********** Public ARRAYS ***********
Public UserList() As User 'Holds data for each user
Public NPCList(1 To MAX_NPCs) As NPC 'Holds data for each NPC
Public MapData() As MapBlock
Public MapInfo() As MapInfo
Public CharList(1 To MAX_CHARACTERS) As Integer
Public ObjData() As ObjData

'********** EXTERNAL FUNCTIONS ***********
'APIs to write and read inis
Declare Function writeprivateprofilestring Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function getprivateprofilestring Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
