Attribute VB_Name = "TCP"
Option Explicit

Public Const SOCKET_BUFFER_SIZE = 20480 'Buffer in bytes for each socket
Public Const COMMAND_BUFFER_SIZE = 1000 'How many commands the server can store from each client

'Constants used in the SendData sub
Public Const ToIndex = 0 'Send data to a single User index
Public Const ToAll = 1 'Send it to all User indexa
Public Const ToMap = 2 'Send it to all users in a map
Public Const ToPCArea = 3 'Send to all users in a user's area
Public Const ToNone = 4 'Send to none
Public Const ToAllButIndex = 5 'Send to all but the index
Public Const ToMapButIndex = 6 'Send to all on a map but the index

' General constants used with most of the controls
Public Const INVALID_HANDLE = -1
Public Const CONTROL_ERRIGNORE = 0
Public Const CONTROL_ERRDISPLAY = 1


' SocketWrench Control Actions
Public Const SOCKET_OPEN = 1
Public Const SOCKET_CONNECT = 2
Public Const SOCKET_LISTEN = 3
Public Const SOCKET_ACCEPT = 4
Public Const SOCKET_CANCEL = 5
Public Const SOCKET_FLUSH = 6
Public Const SOCKET_CLOSE = 7
Public Const SOCKET_DISCONNECT = 7
Public Const SOCKET_ABORT = 8

' SocketWrench Control States
Public Const SOCKET_NONE = 0
Public Const SOCKET_IDLE = 1
Public Const SOCKET_LISTENING = 2
Public Const SOCKET_CONNECTING = 3
Public Const SOCKET_ACCEPTING = 4
Public Const SOCKET_RECEIVING = 5
Public Const SOCKET_SENDING = 6
Public Const SOCKET_CLOSING = 7

' Socket Address Families
Public Const AF_UNSPEC = 0
Public Const AF_UNIX = 1
Public Const AF_INET = 2

' Socket Types
Public Const SOCK_STREAM = 1
Public Const SOCK_DGRAM = 2
Public Const SOCK_RAW = 3
Public Const SOCK_RDM = 4
Public Const SOCK_SEQPACKET = 5

' Protocol Types
Public Const IPPROTO_IP = 0
Public Const IPPROTO_ICMP = 1
Public Const IPPROTO_GGP = 2
Public Const IPPROTO_TCP = 6
Public Const IPPROTO_PUP = 12
Public Const IPPROTO_UDP = 17
Public Const IPPROTO_IDP = 22
Public Const IPPROTO_ND = 77
Public Const IPPROTO_RAW = 255
Public Const IPPROTO_MAX = 256


' Network Addresses
Public Const INADDR_ANY = "0.0.0.0"
Public Const INADDR_LOOPBACK = "127.0.0.1"
Public Const INADDR_NONE = "255.255.255.255"

' Shutdown Values
Public Const SOCKET_READ = 0
Public Const SOCKET_WRITE = 1
Public Const SOCKET_READWRITE = 2

' SocketWrench Error Response
Public Const SOCKET_ERRIGNORE = 0
Public Const SOCKET_ERRDISPLAY = 1

' SocketWrench Error Codes
Public Const WSABASEERR = 24000
Public Const WSAEINTR = 24004
Public Const WSAEBADF = 24009
Public Const WSAEACCES = 24013
Public Const WSAEFAULT = 24014
Public Const WSAEINVAL = 24022
Public Const WSAEMFILE = 24024
Public Const WSAEWOULDBLOCK = 24035
Public Const WSAEINPROGRESS = 24036
Public Const WSAEALREADY = 24037
Public Const WSAENOTSOCK = 24038
Public Const WSAEDESTADDRREQ = 24039
Public Const WSAEMSGSIZE = 24040
Public Const WSAEPROTOTYPE = 24041
Public Const WSAENOPROTOOPT = 24042
Public Const WSAEPROTONOSUPPORT = 24043
Public Const WSAESOCKTNOSUPPORT = 24044
Public Const WSAEOPNOTSUPP = 24045
Public Const WSAEPFNOSUPPORT = 24046
Public Const WSAEAFNOSUPPORT = 24047
Public Const WSAEADDRINUSE = 24048
Public Const WSAEADDRNOTAVAIL = 24049
Public Const WSAENETDOWN = 24050
Public Const WSAENETUNREACH = 24051
Public Const WSAENETRESET = 24052
Public Const WSAECONNABORTED = 24053
Public Const WSAECONNRESET = 24054
Public Const WSAENOBUFS = 24055
Public Const WSAEISCONN = 24056
Public Const WSAENOTCONN = 24057
Public Const WSAESHUTDOWN = 24058
Public Const WSAETOOMANYREFS = 24059
Public Const WSAETIMEDOUT = 24060
Public Const WSAECONNREFUSED = 24061
Public Const WSAELOOP = 24062
Public Const WSAENAMETOOLONG = 24063
Public Const WSAEHOSTDOWN = 24064
Public Const WSAEHOSTUNREACH = 24065
Public Const WSAENOTEMPTY = 24066
Public Const WSAEPROCLIM = 24067
Public Const WSAEUSERS = 24068
Public Const WSAEDQUOT = 24069
Public Const WSAESTALE = 24070
Public Const WSAEREMOTE = 24071
Public Const WSASYSNOTREADY = 24091
Public Const WSAVERNOTSUPPORTED = 24092
Public Const WSANOTINITIALISED = 24093
Public Const WSAHOST_NOT_FOUND = 25001
Public Const WSATRY_AGAIN = 25002
Public Const WSANO_RECOVERY = 25003
Public Const WSANO_DATA = 25004
Public Const WSANO_ADDRESS = 2500

Sub ConnectNewUser(ByVal UserIndex As Integer, ByVal Name As String, ByVal Password As String, ByVal Body As Integer, ByVal Head As Integer)
'*****************************************************************
'Opens a new user. Loads default vars, saves then calls connectuser
'*****************************************************************
Dim LoopC As Integer
  
'Check for Character file
If FileExist(CharPath & UCase(Name) & ".chr", vbNormal) = True Then
    Call SendData(ToIndex, UserIndex, 0, "!!Character already exist.")
    CloseSocket (UserIndex)
    Exit Sub
End If
  
'create file
UserList(UserIndex).Name = Name
UserList(UserIndex).Password = Password
UserList(UserIndex).Char.Heading = SOUTH
UserList(UserIndex).Char.Head = Head
UserList(UserIndex).Char.Body = Body

UserList(UserIndex).Stats.MET = 1
UserList(UserIndex).Stats.MaxHP = 20
UserList(UserIndex).Stats.MinHP = 20

UserList(UserIndex).Stats.FIT = 1
UserList(UserIndex).Stats.MaxSTA = 5
UserList(UserIndex).Stats.MinSTA = 5

UserList(UserIndex).Stats.MaxMAN = 30
UserList(UserIndex).Stats.MinMAN = 30

UserList(UserIndex).Stats.MaxHIT = 2
UserList(UserIndex).Stats.MinHIT = 1

UserList(UserIndex).Stats.GLD = 50

UserList(UserIndex).Stats.EXP = 0
UserList(UserIndex).Stats.ELU = 300
UserList(UserIndex).Stats.ELV = 1

UserList(UserIndex).Object(1).ObjIndex = 1
UserList(UserIndex).Object(1).Amount = 5

Call SaveUser(UserIndex, CharPath & UCase(Name) & ".chr")
  
'Open User
Call ConnectUser(UserIndex, Name, Password)
  
End Sub

Sub CloseSocket(ByVal UserIndex As Integer)
'*****************************************************************
'Close the users socket
'*****************************************************************
On Error Resume Next
  
If UserIndex > 0 Then

    frmMain.Socket2(UserIndex).Disconnect

    If UserList(UserIndex).Flags.UserLogged = 1 Then
        Call CloseUser(UserIndex)
    End If

    UserList(UserIndex).ConnID = -1
    frmMain.Socket2(UserIndex).Cleanup
    Unload frmMain.Socket2(UserIndex)

End If

End Sub


Sub SendData(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal sndData As String)
'*****************************************************************
'Sends data to sendRoute
'*****************************************************************
Dim LoopC As Integer
Dim X As Integer
Dim Y As Integer

'Add End character
sndData = sndData & ENDC
  
'send NONE
If sndRoute = ToNone Then
    Exit Sub
End If
  
  
'Send to All
If sndRoute = ToAll Then
    For LoopC = 1 To LastUser
        If UserList(LoopC).Flags.UserLogged Then
            frmMain.Socket2(LoopC).Write sndData, Len(sndData)
        End If
    Next LoopC
    Exit Sub
End If

'Send to everyone but the sndindex
If sndRoute = ToAllButIndex Then
    For LoopC = 1 To LastUser
              
      If UserList(LoopC).Flags.UserLogged And LoopC <> sndIndex Then
            frmMain.Socket2(LoopC).Write sndData, Len(sndData)
      End If
      
    Next LoopC
    Exit Sub
End If

'Send to Map
If sndRoute = ToMap Then

    For LoopC = 1 To LastUser

        If UserList(LoopC).Flags.UserLogged Then
            If UserList(LoopC).Pos.Map = sndMap Then
                frmMain.Socket2(LoopC).Write sndData, Len(sndData)
            End If
        End If
      
    Next LoopC
    
    Exit Sub
End If

'Send to everone on map but sndIndex
If sndRoute = ToMapButIndex Then

    For LoopC = 1 To LastUser

        If UserList(LoopC).Flags.UserLogged And LoopC <> sndIndex Then
            If UserList(LoopC).Pos.Map = sndMap Then
                frmMain.Socket2(LoopC).Write sndData, Len(sndData)
             End If
        End If
  
    Next LoopC
    
    Exit Sub
End If

'Send to PC Area
If sndRoute = ToPCArea Then
    
    For Y = UserList(sndIndex).Pos.Y - MinYBorder + 1 To UserList(sndIndex).Pos.Y + MinYBorder - 1
        For X = UserList(sndIndex).Pos.X - MinXBorder + 1 To UserList(sndIndex).Pos.X + MinXBorder - 1

            If MapData(sndMap, X, Y).UserIndex > 0 Then

                frmMain.Socket2(MapData(sndMap, X, Y).UserIndex).Write sndData, Len(sndData)

            End If
        
        Next X
    Next Y
    
    Exit Sub
End If

'Send to the UserIndex
If sndRoute = ToIndex Then
    frmMain.Socket2(sndIndex).Write sndData, Len(sndData)
    Exit Sub
End If

End Sub


Sub ConnectUser(ByVal UserIndex As Integer, ByVal Name As String, ByVal Password As String)
'*****************************************************************
'Reads the users .chr file and loads into Userlist array
'*****************************************************************

'Check for max users
If NumUsers >= MaxUsers Then
    Call SendData(ToIndex, UserIndex, 0, "!!Too many users logged on. Try again later.")
    CloseSocket (UserIndex)
    Exit Sub
End If
  
'Check to see is user already logged with IP
If AllowMultiLogins = 0 Then
    If CheckForSameIP(UserIndex, frmMain.Socket2(UserIndex).PeerAddress) = True Then
        Call SendData(ToIndex, UserIndex, 0, "!!Sorry, your IP address is already logged on to the server. Please only use one character at a time.")
        CloseSocket (UserIndex)
        Exit Sub
    End If
End If

'Check to see is user already logged with Name
If CheckForSameName(UserIndex, Name) = True Then
    Call SendData(ToIndex, UserIndex, 0, "!!Sorry, a user with the same name is already logged on.")
    CloseSocket (UserIndex)
    Exit Sub
End If

'Check for Character file
If FileExist(CharPath & UCase(Name) & ".chr", vbNormal) = False Then
    Call SendData(ToIndex, UserIndex, 0, "!!Character does not exist.")
    CloseSocket (UserIndex)
    Exit Sub
End If

'Check Password
If Password <> GetVar(CharPath & UCase(Name) & ".chr", "INIT", "Password") Then
    Call SendData(ToIndex, UserIndex, 0, "!!Wrong Password.")
    CloseSocket (UserIndex)
    Exit Sub
End If

'Load init vars from file
Call LoadUserInit(UserIndex, CharPath & UCase(Name) & ".chr")
Call LoadUserStats(UserIndex, CharPath & UCase(Name) & ".chr")

'Figure out where to put user
If UserList(UserIndex).Pos.Map > 0 Then
    If MapInfo(UserList(UserIndex).Pos.Map).StartPos.Map > 0 Then
        UserList(UserIndex).Pos = MapInfo(UserList(UserIndex).Pos.Map).StartPos
    End If
Else
    UserList(UserIndex).Pos = StartPos
End If

'Get closest legal pos
Call ClosestLegalPos(UserList(UserIndex).Pos, UserList(UserIndex).Pos)
If LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) = False Then
    Call SendData(ToIndex, UserIndex, 0, "!!No legal position found: Please try again.")
    CloseUser (UserIndex)
    Exit Sub
End If

'Get mod name
UserList(UserIndex).Name = Name
If WizCheck(UserList(UserIndex).Name) Then
    UserList(UserIndex).modName = Name & " <GM>"
Else
    UserList(UserIndex).modName = Name
End If

'************** Initialize variables
UserList(UserIndex).Password = Password
UserList(UserIndex).IP = frmMain.Socket2(UserIndex).PeerAddress
UserList(UserIndex).Flags.UserLogged = 1

'Set switching map flag
UserList(UserIndex).Flags.SwitchingMaps = 1
'Send User index
Call SendData(ToIndex, UserIndex, 0, "SUI" & UserIndex)
'Tell client to try switching maps
Call SendData(ToIndex, UserIndex, 0, "SCM" & UserList(UserIndex).Pos.Map & "," & MapInfo(UserList(UserIndex).Pos.Map).MapVersion)
'Welcome message
Call SendData(ToIndex, UserIndex, 0, "@Welcome to the game. For help type /help." & FONTTYPE_INFO)

'Update inventory
Call UpdateUserInv(True, UserIndex, 0)

'update Num of Users
If UserIndex > LastUser Then LastUser = UserIndex
NumUsers = NumUsers + 1
frmMain.txStatus.Text = "Total Users= " & NumUsers
MapInfo(UserList(UserIndex).Pos.Map).NumUsers = MapInfo(UserList(UserIndex).Pos.Map).NumUsers + 1

'Show Character to others
Call MakeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)

'Refresh list box and send log on string
Call RefreshUserListBox
  
'Send login in sound and login phrase
Call SendData(ToIndex, UserIndex, 0, "PLW" & SOUND_WARP)
Call SendData(ToAll, 0, 0, "@" & UserList(UserIndex).Name & " has entered the game." & FONTTYPE_INFO)

'Log it
Open App.Path & "\Connect.log" For Append Shared As #5
Print #5, UserList(UserIndex).Name & " logged in. UserIndex:" & UserIndex & " " & Time & " " & Date
Close #5

End Sub

Sub CloseUser(ByVal UserIndex As Integer)
'*****************************************************************
'save user then reset user's slot
'*****************************************************************
Dim X As Integer
Dim Y As Integer
Dim LoopC As Integer
Dim Map As Integer
Dim Name As String

'Save temps
Map = (UserList(UserIndex).Pos.Map)
X = UserList(UserIndex).Pos.X
Y = UserList(UserIndex).Pos.Y
Name = UserList(UserIndex).Name

'Set logged to false
UserList(UserIndex).Flags.UserLogged = 0

'Save user
Call SaveUser(UserIndex, CharPath & UCase(Name) & ".chr")

'Erase user's character
UserList(UserIndex).Char.Body = 0
UserList(UserIndex).Char.Head = 0
UserList(UserIndex).Char.Heading = 0

If UserList(UserIndex).Char.CharIndex > 0 Then
    Call EraseUserChar(ToMap, 0, Map, UserIndex)
End If

'Clear main vars
UserList(UserIndex).Name = ""
UserList(UserIndex).modName = ""
UserList(UserIndex).Password = ""
UserList(UserIndex).Pos.Map = 0
UserList(UserIndex).Pos.X = 0
UserList(UserIndex).Pos.Y = 0
UserList(UserIndex).IP = ""
UserList(UserIndex).RDBuffer = ""

'Clear Counters
UserList(UserIndex).Counters.IdleCount = 0
UserList(UserIndex).Counters.AttackCounter = 0
UserList(UserIndex).Counters.SendMapCounter.Map = 0
UserList(UserIndex).Counters.SendMapCounter.X = 0
UserList(UserIndex).Counters.SendMapCounter.Y = 0
UserList(UserIndex).Counters.HPCounter = 0
UserList(UserIndex).Counters.STACounter = 0

'Clear Flags
UserList(UserIndex).Flags.DownloadingMap = 0
UserList(UserIndex).Flags.SwitchingMaps = 0
UserList(UserIndex).Flags.StatsChanged = 0
UserList(UserIndex).Flags.ReadyForNextTile = 0

'update last user
If UserIndex = LastUser Then
    Do Until UserList(LastUser).Flags.UserLogged = 1
        LastUser = LastUser - 1
        If LastUser = 0 Then Exit Do
    Loop
End If
  
'update number of users
If NumUsers <> 0 Then
    NumUsers = NumUsers - 1
End If
frmMain.txStatus.Text = "Total Users= " & NumUsers
Call RefreshUserListBox

'Update Map Users
MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1
If MapInfo(Map).NumUsers < 0 Then
    MapInfo(Map).NumUsers = 0
End If

'Send log off phrase
Call SendData(ToAll, 0, 0, "@" & Name & " has left the game." & FONTTYPE_INFO)

'Log it
Open App.Path & "\Connect.log" For Append Shared As #5
Print #5, Name & " logged off. " & "User Index:" & UserIndex & " " & Time & " " & Date
Close #5
  
End Sub

Sub HandleData(ByVal UserIndex As Integer, ByVal Rdata As String)
'*****************************************************************
'Handles all data from the clients
'*****************************************************************
Dim sndData As String
Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim tIndex As Integer
Dim tName As String
Dim tMessage As String
Dim Arg1 As String
Dim Arg2 As String
Dim Arg3 As String
Dim Arg4 As String

'Check to see if user has a valid UserIndex
If UserIndex < 0 Then
    Exit Sub
End If
    
'Reset Idle
UserList(UserIndex).Counters.IdleCount = 0
    
'******************* Login Commands ****************************
    
'Logon on existing character
If Left$(Rdata, 5) = "LOGIN" Then
    Rdata = Right$(Rdata, Len(Rdata) - 5)
    
    Call ConnectUser(UserIndex, ReadField(1, Rdata, 44), ReadField(2, Rdata, 44))

    Exit Sub
End If
  
'Make a new character
If Left$(Rdata, 6) = "NLOGIN" Then
    Rdata = Right$(Rdata, Len(Rdata) - 6)
    
    Call ConnectNewUser(UserIndex, ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), Val(ReadField(3, Rdata, 44)), ReadField(4, Rdata, 44))
    
    Exit Sub
End If
  
'If not trying to log on must not be a client so log it off
If UserList(UserIndex).Flags.UserLogged = 0 Then
    CloseSocket (UserIndex)
    Exit Sub
End If
  
'******************* Communication Commands ****************************
'Say
If Left$(Rdata, 1) = ";" Then
    Rdata = Right$(Rdata, Len(Rdata) - 1)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "@" & UserList(UserIndex).Name & ": " & Rdata & FONTTYPE_TALK)
    Exit Sub
End If

'Broadcast
If Left$(Rdata, 1) = "'" Then
    Rdata = Right$(Rdata, Len(Rdata) - 1)
    
    Call SendData(ToAll, 0, UserList(UserIndex).Pos.Map, "@" & UserList(UserIndex).Name & " yells: " & Rdata & FONTTYPE_TALK)
    Exit Sub
End If
  
'Shout
If Left$(Rdata, 1) = "-" Then
    Rdata = Right$(Rdata, Len(Rdata) - 1)
    Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "@" & UserList(UserIndex).Name & " shouts: " & Rdata & FONTTYPE_TALK)
    Exit Sub
End If
  
'Emote
If Left$(Rdata, 1) = ":" Then
    Rdata = Right$(Rdata, Len(Rdata) - 1)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "@" & UserList(UserIndex).Name & " " & Rdata & FONTTYPE_TALK)
    Exit Sub
End If
  
'Whisper
If Left$(Rdata, 1) = "\" Then
    Rdata = Right$(Rdata, Len(Rdata) - 1)
    
    tName = ReadField(1, Rdata, 32)
    tIndex = NameIndex(tName)
    
    If tIndex <> 0 Then
    
        If Len(Rdata) <> Len(tName) Then
            tMessage = Right$(Rdata, Len(Rdata) - (1 + Len(tName)))
        Else
            tMessage = " "
        End If
        
        Call SendData(ToIndex, tIndex, 0, "@" & UserList(UserIndex).Name & " whispers, " & Chr(34) & tMessage & Chr(34) & " to you." & FONTTYPE_TALK)
        Call SendData(ToIndex, UserIndex, 0, "@You whisper, " & Chr(34) & tMessage & Chr(34) & " to " & UserList(tIndex).Name & "." & FONTTYPE_TALK)
        Exit Sub
    End If
    
    Call SendData(ToIndex, UserIndex, 0, "@User not online. " & FONTTYPE_INFO)
    Exit Sub
End If

'Who
If UCase$(Rdata) = "/WHO" Then
    Call SendData(ToIndex, UserIndex, 0, "@Total Users: " & NumUsers & FONTTYPE_INFO)
    
    For LoopC = 1 To LastUser
        If (UserList(LoopC).Name <> "") Then
            tStr = tStr & UserList(LoopC).modName & ", "
        End If
    Next LoopC
    tStr = Left$(tStr, Len(tStr) - 2)
    
    Call SendData(ToIndex, UserIndex, 0, "@" & tStr & FONTTYPE_INFO)
    
    Exit Sub
End If


'******************* General Commands ****************************

'Move
If Left$(Rdata, 1) = "M" Then
    'Don't allow if switching maps
    If UserList(UserIndex).Flags.SwitchingMaps Then
        Exit Sub
    End If
    Rdata = Right$(Rdata, Len(Rdata) - 1)
    Call MoveUserChar(UserIndex, Val(Rdata))
    Exit Sub
End If

'Rotate Right
If Rdata = ">" Then
    'Don't allow if switching maps maps
    If UserList(UserIndex).Flags.SwitchingMaps Then
        Exit Sub
    End If
    UserList(UserIndex).Char.Heading = UserList(UserIndex).Char.Heading + 1
    If UserList(UserIndex).Char.Heading > WEST Then UserList(UserIndex).Char.Heading = NORTH
    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading)
    Exit Sub
End If
  
'Rotate Left
If Rdata = "<" Then
    'Don't allow if switching maps
    If UserList(UserIndex).Flags.SwitchingMaps Then
        Exit Sub
    End If
    UserList(UserIndex).Char.Heading = UserList(UserIndex).Char.Heading - 1
    If UserList(UserIndex).Char.Heading < NORTH Then UserList(UserIndex).Char.Heading = WEST
    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading)
End If

'Attack
If Rdata = "ATT" Then
    'Don't allow if switching maps
    If UserList(UserIndex).Flags.SwitchingMaps Then
        Exit Sub
    End If
    UserAttack UserIndex
    Exit Sub
End If

'Left Click
If Left$(Rdata, 2) = "LC" Then
    'Don't allow if switching maps
    If UserList(UserIndex).Flags.SwitchingMaps Then
        Call SendData(ToIndex, UserIndex, 0, "@User switching maps." & FONTTYPE_INFO)
        Exit Sub
    End If
    Rdata = Right$(Rdata, Len(Rdata) - 2)
    Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, ReadField(1, Rdata, 44), ReadField(2, Rdata, 44))
    Exit Sub
End If

'Right Click
If Left$(Rdata, 2) = "RC" Then
    'Don't allow if switching maps
    If UserList(UserIndex).Flags.SwitchingMaps Then
        Call SendData(ToIndex, UserIndex, 0, "@User switching maps." & FONTTYPE_INFO)
        Exit Sub
    End If
    Exit Sub
End If

'HELP
If UCase$(Rdata) = "/HELP" Then
    Call SendHelp(UserIndex)
    Exit Sub
End If

'Quit
If UCase$(Rdata) = "/QUIT" Then
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'******************* Map Commands ****************************

'Request Map Update
If Left$(Rdata, 3) = "RMU" Then
    Rdata = Right$(Rdata, Len(Rdata) - 3)
    
    UserList(UserIndex).Flags.DownloadingMap = 1
    UserList(UserIndex).Flags.ReadyForNextTile = 1
    UserList(UserIndex).Counters.SendMapCounter.Map = Val(Rdata)
    UserList(UserIndex).Counters.SendMapCounter.X = XMinMapSize
    UserList(UserIndex).Counters.SendMapCounter.Y = YMinMapSize
    
    Call SendData(ToIndex, UserIndex, 0, "SMT" & MapInfo(Val(Rdata)).MapVersion)

End If

'Request Pos update
If Rdata = "RPU" Then
    Call SendData(ToIndex, UserIndex, 0, "SUP" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
    Exit Sub
End If

'Ready for next tile
If Rdata = "RNT" Then
    UserList(UserIndex).Flags.ReadyForNextTile = 1
    Exit Sub
End If

'Done Loading Map
If Rdata = "DLM" Then
    UserList(UserIndex).Flags.SwitchingMaps = 0
    Call SendData(ToIndex, UserIndex, 0, "SMN" & MapInfo(UserList(UserIndex).Pos.Map).Name)
    Call SendData(ToIndex, UserIndex, 0, "PLM" & MapInfo(UserList(UserIndex).Pos.Map).Music)
    
    Call SendData(ToIndex, UserIndex, 0, "DSM") 'Tell client to start drawing
    
    Call UpdateUserMap(UserIndex) 'Fill in all the characters and objects
    Call SendData(ToIndex, UserIndex, 0, "SUC" & UserList(UserIndex).Char.CharIndex)
    
    Exit Sub
End If

'******************* Object Commands ****************************

'Get
If Rdata = "GET" Then
    'Don't allow if switching maps
    If UserList(UserIndex).Flags.SwitchingMaps Then
        Exit Sub
    End If
    Call GetObj(UserIndex)
    Exit Sub
End If
  
'Drop
If Left$(Rdata, 3) = "DRP" Then
    'Don't allow if switching maps
    If UserList(UserIndex).Flags.SwitchingMaps Then
        Exit Sub
    End If
    Rdata = Right$(Rdata, Len(Rdata) - 3)
    If UserList(UserIndex).Object(ReadField(1, Rdata, 44)).ObjIndex = 0 Then
        Exit Sub
    End If
    Call DropObj(UserIndex, Val(ReadField(1, Rdata, 44)), Val(ReadField(2, Rdata, 44)), UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    Exit Sub
End If

'USE
If Left$(Rdata, 3) = "USE" Then
    'Don't allow if switching maps
    If UserList(UserIndex).Flags.SwitchingMaps Then
        Exit Sub
    End If
    Rdata = Right$(Rdata, Len(Rdata) - 3)
    If UserList(UserIndex).Object(Val(Rdata)).ObjIndex = 0 Then
        Exit Sub
    End If
    Call UseInvItem(UserIndex, Val(Rdata))
    Exit Sub
End If

'******************* Status Commands ****************************

'Stats
If UCase(Rdata) = "/STATS" Then
    SendUserStatsTxt UserIndex, UserIndex
    Exit Sub
End If

'Save
If UCase$(Rdata) = "/SAVE" Then
    Call SaveUser(UserIndex, CharPath & UCase(UserList(UserIndex).Name) & ".chr")
    Call SendData(ToIndex, UserIndex, 0, "@Character saved." & FONTTYPE_INFO)
    Exit Sub
End If

'Change Desc
If UCase$(Left$(Rdata, 6)) = "/DESC " Then
    Rdata = Right$(Rdata, Len(Rdata) - 6)
    UserList(UserIndex).Desc = Rdata
    Call SendData(ToIndex, UserIndex, 0, "@Description changed." & FONTTYPE_INFO)
    Exit Sub
End If

'*************** Wizard commands *****************************
If WizCheck(UserList(UserIndex).Name) = False Then
    Exit Sub
End If

'Reset Server
If UCase$(Rdata) = "/RESET" Then
    
    'Log it
    Open App.Path & "\Main.log" For Append Shared As #5
    Print #5, "!Reset started by " & UserList(UserIndex).Name & ". " & Time & " " & Date
    Close #5
    
    Call Restart
    Exit Sub
End If
  
'Shutdown server
If UCase$(Rdata) = "/SHUTDOWN" Then
    'Log it
    Open App.Path & "\Main.log" For Append Shared As #5
    Print #5, "!Shutdown started by " & UserList(UserIndex).Name & ". " & Time & " " & Date
    Close #5
    
    Unload frmMain
    Exit Sub
End If

'System Message
If UCase$(Left$(Rdata, 6)) = "/SMSG " Then
    Rdata = Right$(Rdata, Len(Rdata) - 6)
    
    If Rdata <> "" Then
        Call SendData(ToAll, 0, 0, "!" & Rdata)
    End If
    
    Exit Sub
End If

'Spoof
If UCase$(Left$(Rdata, 7)) = "/SPOOF " Then
    Rdata = Right$(Rdata, Len(Rdata) - 7)
    
    tIndex = NameIndex(ReadField(1, Rdata, 32))
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "@User not online." & FONTTYPE_INFO)
        Exit Sub
    End If
    Call SendData(ToPCArea, tIndex, UserList(tIndex).Pos.Map, "@" & Rdata & FONTTYPE_TALK)
    
    Exit Sub
End If
  
'Emergency System Message
If UCase$(Left$(Rdata, 6)) = "/EMSG " Then
    Rdata = Right$(Rdata, Len(Rdata) - 6)
    
    If Rdata <> "" Then
        Call SendData(ToAll, 0, 0, "!!" & Rdata)
    End If
    
    Exit Sub
End If

'Control Code (send a command to all the clients)
If UCase$(Left$(Rdata, 4)) = "/CC " Then
    Rdata = Right$(Rdata, Len(Rdata) - 4)
    
    If Rdata <> "" Then
        Call SendData(ToAll, 0, 0, Rdata)
    End If
    
    Exit Sub
End If

'RP Message
If UCase$(Left$(Rdata, 6)) = "/RMSG " Then
    Rdata = Right$(Rdata, Len(Rdata) - 6)
    
    If Rdata <> "" Then
        Call SendData(ToAll, 0, 0, "@" & Rdata & FONTTYPE_TALK)
    End If
    
    Exit Sub
End If

'Time
If UCase$(Left$(Rdata, 5)) = "/TIME" Then
    Rdata = Right$(Rdata, Len(Rdata) - 5)
    
        Call SendData(ToAll, 0, 0, "@At the tone, the server time will be: " & Time & " " & Date & FONTTYPE_INFO)
    
    Exit Sub
End If

'Where is
If UCase$(Left$(Rdata, 9)) = "/WHEREIS " Then
    Rdata = Right$(Rdata, Len(Rdata) - 9)
    
    tIndex = NameIndex(Rdata)
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "@User not online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call SendData(ToIndex, UserIndex, 0, "@Loc for " & UserList(tIndex).Name & ": " & UserList(tIndex).Pos.Map & ", " & UserList(tIndex).Pos.X & ", " & UserList(tIndex).Pos.Y & "." & FONTTYPE_INFO)
    
    Exit Sub
End If

'Approach
If UCase$(Left$(Rdata, 6)) = "/APPR " Then
    Rdata = Right$(Rdata, Len(Rdata) - 6)

    'Don't allow if switching maps
    If UserList(UserIndex).Flags.SwitchingMaps Then
        Call SendData(ToIndex, UserIndex, 0, "@User switching maps." & FONTTYPE_INFO)
        Exit Sub
    End If

    'See if user online
    tIndex = NameIndex(Rdata)
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "@User not online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    'Find closest legal position and warp there
    ClosestLegalPos UserList(tIndex).Pos, nPos
    If LegalPos(nPos.Map, nPos.X, nPos.Y) Then
        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)
        Call SendData(ToIndex, tIndex, 0, "@" & UserList(UserIndex).Name & " approached you." & FONTTYPE_INFO)
    End If
    
    Exit Sub
End If

'Summon
If UCase$(Left$(Rdata, 5)) = "/SUM " Then
    Rdata = Right$(Rdata, Len(Rdata) - 5)
    
    'See if user online
    tIndex = NameIndex(Rdata)
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "@User not online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    'Don't allow if switching maps
    If UserList(tIndex).Flags.SwitchingMaps Then
        Call SendData(ToIndex, UserIndex, 0, "@User switching maps." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    'Find closest legal position and warp there
    ClosestLegalPos UserList(UserIndex).Pos, nPos
    If LegalPos(nPos.Map, nPos.X, nPos.Y) Then
        Call SendData(ToIndex, tIndex, 0, "@" & UserList(UserIndex).Name & " has summoned you." & FONTTYPE_INFO)
        Call WarpUserChar(tIndex, nPos.Map, nPos.X, nPos.Y)
    End If
    
    Exit Sub
End If

'GM Message
If UCase$(Left$(Rdata, 6)) = "/GMSG " Then
    Rdata = Right$(Rdata, Len(Rdata) - 6)
    
    If Rdata <> "" Then
        Call SendData(ToAll, 0, 0, "@" & UserList(UserIndex).Name & "!>" & Rdata & FONTTYPE_INFO)
    End If
    
    Exit Sub
End If

'Boot user
If UCase$(Left$(Rdata, 6)) = "/BOOT " Then
    Rdata = Right$(Rdata, Len(Rdata) - 6)
    
    tIndex = NameIndex(Rdata)
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "@User not online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    'Log it
    Open App.Path & "\Main.log" For Append Shared As #5
    Print #5, "" & UserList(UserIndex).Name & " booted " & UserList(tIndex).Name & ". " & Time & " " & Date
    Close #5
    
    Call SendData(ToAll, 0, 0, "@" & UserList(UserIndex).Name & " booted " & UserList(tIndex).Name & "." & FONTTYPE_INFO)
    CloseSocket (tIndex)
    
    Exit Sub
End If

'Character modify
If UCase(Left(Rdata, 9)) = "/CHARMOD " Then
    Rdata = Right$(Rdata, Len(Rdata) - 9)
    
    tIndex = NameIndex(ReadField(1, Rdata, 32))
    Arg1 = ReadField(2, Rdata, 32)
    Arg2 = ReadField(3, Rdata, 32)
    Arg3 = ReadField(4, Rdata, 32)
    Arg4 = ReadField(5, Rdata, 32)

    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "@User not online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    'Don't allow if switching maps maps
    If UserList(tIndex).Flags.SwitchingMaps Then
        Call SendData(ToIndex, UserIndex, 0, "@User switching maps." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Select Case UCase(Arg1)
    
        Case "GLD"
            UserList(tIndex).Stats.GLD = Val(Arg2)
            Call SendUserStatsBox(tIndex)

        Case "LVL"
            UserList(tIndex).Stats.ELV = Val(Arg2)
            Call SendUserStatsBox(tIndex)
    
        Case "BODY"
            Call ChangeUserChar(ToMap, 0, UserList(tIndex).Pos.Map, tIndex, Val(Arg2), UserList(tIndex).Char.Head, UserList(tIndex).Char.Heading)

        Case "HEAD"
            Call ChangeUserChar(ToMap, 0, UserList(tIndex).Pos.Map, tIndex, UserList(tIndex).Char.Body, Val(Arg2), UserList(tIndex).Char.Heading)
        
        Case "WARP"
            If LegalPos(Val(Arg2), Val(Arg3), Val(Arg4)) Then
                Call WarpUserChar(tIndex, Val(Arg2), Val(Arg3), Val(Arg4))
            Else
                Call SendData(ToIndex, UserIndex, 0, "@Not a legal position." & FONTTYPE_INFO)
            End If
        
        Case Else
            Call SendData(ToIndex, UserIndex, 0, "@Not a charmod command." & FONTTYPE_INFO)
    
    End Select

    Exit Sub
End If

'**************************************

End Sub



