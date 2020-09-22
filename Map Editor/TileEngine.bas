Attribute VB_Name = "TileEngine"
Option Explicit
'*************************************************************
'TileEngine v 1.3  9-4-00
'Copyrighted 2000 Baronsoft
'
'Aaron Perkins
'aaron@baronsoft.com
'http://www.baronsoft.com
'*************************************************************

'********** CONSTANTS ***********
'Heading Constants
Public Const NORTH = 1
Public Const EAST = 2
Public Const SOUTH = 3
Public Const WEST = 4

'Map sizes in tiles
Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

'Object Constants
Public Const MAX_INVENORY_OBJS = 99

'bltbit constant
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

'Sound flag constants
Public Const SND_SYNC = &H0 ' play synchronously (default)
Public Const SND_ASYNC = &H1 ' play asynchronously
Public Const SND_NODEFAULT = &H2 ' silence not default, if sound not found
Public Const SND_LOOP = &H8 ' loop the sound until next sndPlaySound
Public Const SND_NOSTOP = &H10 ' don't stop any currently playing sound

Public Const NumSoundBuffers = 7

'********** TYPES ***********

'Bitmap header
Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

'Bitmap info header
Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

'Holds a local position
Public Type Position
    X As Integer
    Y As Integer
End Type

'Holds a world position
Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

'Holds data about where a bmp can be found,
'How big it is and animation info
Public Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames(1 To 16) As Integer
    Speed As Integer
End Type

'Points to a grhData and keeps animation info
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Byte
    SpeedCounter As Byte
    Started As Byte
End Type

'Bodies list
Public Type BodyData
    Walk(1 To 4) As Grh
    HeadOffset As Position
End Type

'Heads list
Public Type HeadData
    Head(1 To 4) As Grh
End Type

'Hold info about a character
Public Type Char
    Active As Byte
    Heading As Byte
    Pos As Position

    Body As BodyData
    Head As HeadData
    
    Moving As Byte
    MoveOffset As Position
End Type

'Holds info about a object
Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

'Holds info about each tile position
Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
End Type

'Hold info about each map
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    
    'ME Only
    Changed As Byte
End Type

'********** Public VARS ***********
'Paths
Public GrhPath As String
Public IniPath As String
Public MapPath As String

'Where the map borders are.. Set during load
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'User status vars
Public CurMap As Integer 'Current map loaded
Public UserIndex As Integer
Public UserMoving As Byte
Global UserBody As Integer
Global UserHead As Integer
Public UserPos As Position 'Holds current user pos
Public AddtoUserPos As Position 'For moving user
Public UserCharIndex As Integer

Public EngineRun As Boolean
Public FramesPerSec As Integer
Public FramesPerSecCounter As Long

'Main view size size in tiles
Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

'Pixel offset of main view screen from 0,0
Public MainViewTop As Integer
Public MainViewLeft As Integer

'How many tiles the engine "looks ahead" when
'drawing the screen
Public TileBufferSize As Integer

'Handle to where all the drawing is going to take place
Public DisplayFormhWnd As Long

'Tile size in pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

'Map editor variables
Public WalkMode As Boolean
Public DrawGrid As Boolean
Public DrawBlock As Boolean

'Totals
Public NumMaps As Integer 'Number of maps
Public NumBodies As Integer
Public NumHeads As Integer
Public NumGrhFiles As Integer 'Number of bmps
Public NumGrhs As Integer 'Number of Grhs
Global NumChars As Integer
Global LastChar As Integer

'********** Direct X ***********
Public MainViewRect As RECT
Public MainViewWidth As Integer
Public MainViewHeight As Integer
Public BackBufferRect As RECT

Public DirectX As New DirectX7
Public DirectDraw As DirectDraw7

Public PrimarySurface As DirectDrawSurface7
Public PrimaryClipper As DirectDrawClipper
Public BackBufferSurface As DirectDrawSurface7
Public SurfaceDB() As DirectDrawSurface7

'Sound
Dim DirectSound As DirectSound
Dim DSBuffers(1 To NumSoundBuffers) As DirectSoundBuffer
Dim LastSoundBufferUsed As Integer

'********** Public ARRAYS ***********
Public GrhData() As GrhData 'Holds all the grh data

Public BodyData() As BodyData
Public HeadData() As HeadData

Public MapData() As MapBlock 'Holds map data for current map
Public MapInfo As MapInfo 'Holds map info for current map
Public CharList(1 To 10000) As Char 'Holds info about all characters on map


'********** OUTSIDE FUNCTIONS ***********
'Good old BitBlt
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'Sound stuff
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uRetrunLength As Long, ByVal hwndCallback As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'Sleep
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Static Sub EvenFrameSpeed(CurrentFPS As Integer, TargetFPS As Integer)
'*****************************************************************
'Delays the engine attempting to keep the FPS close to TargetFPS
'*****************************************************************
    Dim FrameCounter As Integer
    Dim OffsetFPS As Integer
    Dim Delay As Integer

    FrameCounter = FrameCounter + 1

    'Recalculate delay every TragetFPS frames
    If FrameCounter >= TargetFPS Then
        FrameCounter = 0
        
        OffsetFPS = CurrentFPS - TargetFPS
        
        If OffsetFPS > 0 Then
            Delay = OffsetFPS * (TargetFPS / CurrentFPS)
        End If
    End If

    'Delay
    Sleep Delay

End Sub


Function LoadWavetoDSBuffer(DS As DirectSound, DSB As DirectSoundBuffer, sfile As String) As Boolean


    '========================================================================
    '- Step4 CREATE SOUND BUFFER FROM FILE.
    '  we use the DSBUFFERDESC type to indicate
    '  what features we want the sound to have.
    '  The lFlags member can be used to enable 3d support,
    '  frequency changes, and volume changes.
    '  The DSBCAPS flags indicates we will allow
    '  volume changes, frequency changes, and pan changes
    '  the DDSBCAPS_STATIC -(which is optional in this release
    '  since all  buffers loaded by this method are static) indicates
    '  that we want the entire file loaded into memory.
    '
    '  The function fills in the other members of bufferDesc which lets
    '  us know how large the buffer is.  It also fills in the wave Format
    '  type giving information about the waves quality and if it supports
    '  stereo the function returns an initialized SoundBuffer
    '=========================================================================
    
    Dim bufferDesc As DSBUFFERDESC
    Dim waveFormat As WAVEFORMATEX
    
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2
    waveFormat.lSamplesPerSec = 22050
    waveFormat.nBitsPerSample = 16
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    Set DSB = DS.CreateSoundBufferFromFile(sfile, bufferDesc, waveFormat)
    
    '========================================
    '- Step 5 make sure we have no errors
    '========================================
    
    If Err.Number <> 0 Then
        'MsgBox "unable to find " + sfile
        'End
        Exit Function
    End If
    
    LoadWavetoDSBuffer = True
    
End Function
Sub LoadHeadData()
'*****************************************************************
'Loads Head.dat
'*****************************************************************

Dim LoopC As Integer

'Get Number of heads
NumHeads = Val(GetVar(IniPath & "Head.dat", "INIT", "NumHeads"))

'Resize array
ReDim HeadData(1 To NumHeads) As HeadData

'Fill List
For LoopC = 1 To NumHeads
    InitGrh HeadData(LoopC).Head(1), Val(GetVar(IniPath & "Head.dat", "Head" & LoopC, "Head1")), 0
    InitGrh HeadData(LoopC).Head(2), Val(GetVar(IniPath & "Head.dat", "Head" & LoopC, "Head2")), 0
    InitGrh HeadData(LoopC).Head(3), Val(GetVar(IniPath & "Head.dat", "Head" & LoopC, "Head3")), 0
    InitGrh HeadData(LoopC).Head(4), Val(GetVar(IniPath & "Head.dat", "Head" & LoopC, "Head4")), 0
Next LoopC

End Sub

Sub LoadBodyData()
'*****************************************************************
'Loads Body.dat
'*****************************************************************

Dim LoopC As Integer

'Get number of bodies
NumBodies = Val(GetVar(IniPath & "Body.dat", "INIT", "NumBodies"))

'Resize array
ReDim BodyData(1 To NumBodies) As BodyData

'Fill list
For LoopC = 1 To NumBodies
    InitGrh BodyData(LoopC).Walk(1), Val(GetVar(IniPath & "Body.dat", "Body" & LoopC, "Walk1")), 0
    InitGrh BodyData(LoopC).Walk(2), Val(GetVar(IniPath & "Body.dat", "Body" & LoopC, "Walk2")), 0
    InitGrh BodyData(LoopC).Walk(3), Val(GetVar(IniPath & "Body.dat", "Body" & LoopC, "Walk3")), 0
    InitGrh BodyData(LoopC).Walk(4), Val(GetVar(IniPath & "Body.dat", "Body" & LoopC, "Walk4")), 0

    BodyData(LoopC).HeadOffset.X = Val(GetVar(IniPath & "Body.dat", "Body" & LoopC, "HeadOffsetX"))
    BodyData(LoopC).HeadOffset.Y = Val(GetVar(IniPath & "Body.dat", "Body" & LoopC, "HeadOffsetY"))

Next LoopC

End Sub
Sub ConvertCPtoTP(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal CX As Single, ByVal CY As Single, tX As Integer, tY As Integer)
'******************************************
'Converts where the user clicks in the main window
'to a tile position
'******************************************
Dim HWindowX As Integer
Dim HWindowY As Integer

CX = CX - StartPixelLeft
CY = CY - StartPixelTop

HWindowX = (WindowTileWidth \ 2)
HWindowY = (WindowTileHeight \ 2)

'Figure out X and Y tiles
CX = (CX \ TilePixelWidth)
CY = (CY \ TilePixelHeight)

If CX > HWindowX Then
    CX = (CX - HWindowX)

Else
    If CX < HWindowX Then
        CX = (0 - (HWindowX - CX))
    Else
        CX = 0
    End If
End If

If CY > HWindowY Then
    CY = (0 - (HWindowY - CY))
Else
    If CY < HWindowY Then
        CY = (CY - HWindowY)
    Else
        CY = 0
    End If
End If

tX = UserPos.X + CX
tY = UserPos.Y + CY

End Sub




Function DeInitTileEngine() As Boolean
'*****************************************************************
'Shutsdown engine
'*****************************************************************
Dim LoopC As Integer

EngineRun = False

'****** Clear DirectX objects ******
Set PrimarySurface = Nothing
Set PrimaryClipper = Nothing
Set BackBufferSurface = Nothing

'Clear GRH memory
For LoopC = 1 To NumGrhFiles
    Set SurfaceDB(LoopC) = Nothing
Next LoopC
Set DirectDraw = Nothing

'Reset any channels that are done
For LoopC = 1 To NumSoundBuffers
    Set DSBuffers(LoopC) = Nothing
Next LoopC
Set DirectSound = Nothing

Set DirectX = Nothing

DeInitTileEngine = True

End Function

Sub MakeChar(CharIndex As Integer, Body As Integer, Head As Integer, Heading As Byte, X As Integer, Y As Integer)
'*****************************************************************
'Makes a new character and puts it on the map
'*****************************************************************

'Update LastChar
If CharIndex > LastChar Then LastChar = CharIndex
NumChars = NumChars + 1

'Update head, body, ect.
CharList(CharIndex).Body = BodyData(Body)
CharList(CharIndex).Head = HeadData(Head)
CharList(CharIndex).Heading = Heading

'Reset moving stats
CharList(CharIndex).Moving = 0
CharList(CharIndex).MoveOffset.X = 0
CharList(CharIndex).MoveOffset.Y = 0

'Update position
CharList(CharIndex).Pos.X = X
CharList(CharIndex).Pos.Y = Y

'Make active
CharList(CharIndex).Active = 1

'Plot on map
MapData(X, Y).CharIndex = CharIndex

End Sub



Sub EraseChar(CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************

'Make un-active
CharList(CharIndex).Active = 0

'Update lastchar
If CharIndex = LastChar Then
    Do Until CharList(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If

'Remove from map
MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).CharIndex = 0

'Update NumChars
NumChars = NumChars - 1

End Sub

Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************

Grh.GrhIndex = GrhIndex

If Started = 2 Then
    If GrhData(Grh.GrhIndex).NumFrames > 1 Then
        Grh.Started = 1
    Else
        Grh.Started = 0
    End If
Else
    Grh.Started = Started
End If

Grh.FrameCounter = 1
Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed

End Sub

Sub MoveCharbyHead(CharIndex As Integer, nHeading As Byte)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
Dim addX As Integer
Dim addY As Integer
Dim X As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

X = CharList(CharIndex).Pos.X
Y = CharList(CharIndex).Pos.Y

'Figure out which way to move
Select Case nHeading

    Case NORTH
        addY = -1

    Case EAST
        addX = 1

    Case SOUTH
        addY = 1
    
    Case WEST
        addX = -1
        
End Select

nX = X + addX
nY = Y + addY

MapData(nX, nY).CharIndex = CharIndex
CharList(CharIndex).Pos.X = nX
CharList(CharIndex).Pos.Y = nY
MapData(X, Y).CharIndex = 0

CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nHeading

End Sub

Sub MoveCharbyPos(CharIndex As Integer, nX As Integer, nY As Integer)
'*****************************************************************
'Starts the movement of a character to nX,nY
'*****************************************************************
Dim X As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nHeading As Byte

X = CharList(CharIndex).Pos.X
Y = CharList(CharIndex).Pos.Y

addX = nX - X
addY = nY - Y

If Sgn(addX) = 1 Then
    nHeading = EAST
End If

If Sgn(addX) = -1 Then
    nHeading = WEST
End If

If Sgn(addY) = -1 Then
    nHeading = NORTH
End If

If Sgn(addY) = 1 Then
    nHeading = SOUTH
End If

MapData(nX, nY).CharIndex = CharIndex
CharList(CharIndex).Pos.X = nX
CharList(CharIndex).Pos.Y = nY
MapData(X, Y).CharIndex = 0

CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nHeading

End Sub

Sub MoveScreen(Heading As Byte)
'******************************************
'Starts the screen moving in a direction
'******************************************
Dim X As Integer
Dim Y As Integer
Dim tX As Integer
Dim tY As Integer

'Figure out which way to move
Select Case Heading

    Case NORTH
        Y = -1

    Case EAST
        X = 1

    Case SOUTH
        Y = 1
    
    Case WEST
        X = -1
        
End Select

'Fill temp pos
tX = UserPos.X + X
tY = UserPos.Y + Y

'Check to see if its out of bounds
If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
    Exit Sub
Else
    'Start moving... MainLoop does the rest
    AddtoUserPos.X = X
    UserPos.X = tX
    AddtoUserPos.Y = Y
    UserPos.Y = tY
    UserMoving = 1
End If

End Sub


Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
Dim LoopC As Integer

LoopC = 1
Do While CharList(LoopC).Active
    LoopC = LoopC + 1
Loop

NextOpenChar = LoopC

End Function

Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************

Dim LoopC As Integer

For LoopC = 1 To LastChar
    If CharList(LoopC).Active = 1 Then
        MapData(CharList(LoopC).Pos.X, CharList(LoopC).Pos.Y).CharIndex = LoopC
    End If
Next LoopC
    
End Sub
Sub LoadGrhData()
'*****************************************************************
'Loads Grh.dat
'*****************************************************************

On Error GoTo ErrorHandler

Dim Grh As Integer
Dim Frame As Integer
Dim TempInt As Integer

'Get Number of Graphics
GrhPath = GetVar(IniPath & "Grh.ini", "INIT", "Path")
NumGrhs = Val(GetVar(IniPath & "Grh.ini", "INIT", "NumGrhs"))

'Resize arrays
ReDim GrhData(1 To NumGrhs) As GrhData

'Open files
Open IniPath & "Grh.dat" For Binary As #1
Seek #1, 1

'Get Header
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt

'Fill Grh List

'Get first Grh Number
Get #1, , Grh

Do Until Grh <= 0
        
    'Get number of frames
    Get #1, , GrhData(Grh).NumFrames
    If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
    
    If GrhData(Grh).NumFrames > 1 Then
    
        'Read a animation GRH set
        For Frame = 1 To GrhData(Grh).NumFrames
        
            Get #1, , GrhData(Grh).Frames(Frame)
            If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > NumGrhs Then GoTo ErrorHandler
        
        Next Frame
    
        Get #1, , GrhData(Grh).Speed
        If GrhData(Grh).Speed <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
        If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
        If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler
    
    Else
    
        'Read in normal GRH data
        Get #1, , GrhData(Grh).FileNum
        If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sX
        If GrhData(Grh).sX < 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sY
        If GrhData(Grh).sY < 0 Then GoTo ErrorHandler
            
        Get #1, , GrhData(Grh).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / TilePixelHeight
        GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / TilePixelWidth
        
        GrhData(Grh).Frames(1) = Grh
            
    End If

    'Get Next Grh Number
    Get #1, , Grh

Loop
'************************************************

Close #1

Exit Sub

ErrorHandler:
Close #1
MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh

End Sub

Function LegalPos(X As Integer, Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************

'Check to see if its out of bounds
If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    LegalPos = False
    Exit Function
End If

'Check to see if its blocked
If MapData(X, Y).Blocked = 1 Then
    LegalPos = False
    Exit Function
End If

'Check for character
If MapData(X, Y).CharIndex > 0 Then
    LegalPos = False
    Exit Function
End If

LegalPos = True

End Function




Function InMapLegalBounds(X As Integer, Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps
'LEGAL/Walkable bounds
'*****************************************************************

If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

Function InMapBounds(X As Integer, Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************

If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function
Sub DDrawGrhtoSurface(Surface As DirectDrawSurface7, Grh As Grh, X As Integer, Y As Integer, Center As Byte, Animate As Byte)
'*****************************************************************
'Draws a Grh at the X and Y positions
'*****************************************************************
Dim CurrentGrh As Grh
Dim DestRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

'Check to make sure it is legal
If Grh.GrhIndex < 1 Then
    Exit Sub
End If
If GrhData(Grh.GrhIndex).NumFrames < 1 Then
    Exit Sub
End If

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                End If
            End If
        End If
    End If
End If

'Figure out what frame to draw (always 1 if not animated)
CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If Center Then
    If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * TilePixelWidth / 2) + TilePixelWidth / 2
    End If
    If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
    End If
End If

With DestRect
    .Left = X
    .Top = Y
    .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
    .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
End With
    
Surface.GetSurfaceDesc SurfaceDesc

'Draw

If DestRect.Left >= 0 And DestRect.Top >= 0 And DestRect.Right <= SurfaceDesc.lWidth And DestRect.Bottom <= SurfaceDesc.lHeight Then
    
    With SourceRect
        .Left = GrhData(CurrentGrh.GrhIndex).sX
        .Top = GrhData(CurrentGrh.GrhIndex).sY
        .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
    End With
    
    Surface.BltFast DestRect.Left, DestRect.Top, SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum), SourceRect, DDBLTFAST_WAIT
    
End If

End Sub

Sub DDrawTransGrhtoSurface(Surface As DirectDrawSurface7, Grh As Grh, X As Integer, Y As Integer, Center As Byte, Animate As Byte)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
Dim CurrentGrh As Grh
Dim DestRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

'Check to make sure it is legal
If Grh.GrhIndex < 1 Then
    Exit Sub
End If
If GrhData(Grh.GrhIndex).NumFrames < 1 Then
    Exit Sub
End If

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                End If
            End If
        End If
    End If
End If

'Figure out what frame to draw (always 1 if not animated)
CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If Center Then
    If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * TilePixelWidth / 2) + TilePixelWidth / 2
    End If
    If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
    End If
End If

With DestRect
    .Left = X
    .Top = Y
    .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
    .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
End With

Surface.GetSurfaceDesc SurfaceDesc

'Draw
If DestRect.Left >= 0 And DestRect.Top >= 0 And DestRect.Right <= SurfaceDesc.lWidth And DestRect.Bottom <= SurfaceDesc.lHeight Then
    With SourceRect
        .Left = GrhData(CurrentGrh.GrhIndex).sX
        .Top = GrhData(CurrentGrh.GrhIndex).sY
        .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
    End With
    
    Surface.BltFast DestRect.Left, DestRect.Top, SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If

End Sub

Sub DrawBackBufferSurface()
'*****************************************************************
'Copies backbuffer to primarysurface
'*****************************************************************
Dim SourceRect As RECT

With SourceRect
    .Left = (TilePixelWidth * TileBufferSize) - TilePixelWidth
    .Top = (TilePixelHeight * TileBufferSize) - TilePixelHeight
    .Right = .Left + MainViewWidth
    .Bottom = .Top + MainViewHeight
End With

PrimarySurface.Blt MainViewRect, BackBufferSurface, SourceRect, DDBLT_WAIT
'PrimarySurface.Flip Nothing, DDFLIP_WAIT

End Sub

Function GetBitmapDimensions(BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'*****************************************************************
'Gets the dimensions of a bmp
'*****************************************************************
Dim BMHeader As BITMAPFILEHEADER
Dim BINFOHeader As BITMAPINFOHEADER

Open BmpFile For Binary Access Read As #1
Get #1, , BMHeader
Get #1, , BINFOHeader
Close #1

bmWidth = BINFOHeader.biWidth
bmHeight = BINFOHeader.biHeight

End Function



Sub DrawGrhtoHdc(DestHdc As Long, Grh As Grh, X As Integer, Y As Integer, Center As Byte, Animate As Byte, ROP As Long)
'*****************************************************************
'Draws a Grh at the X and Y positions
'*****************************************************************
Dim retcode As Long
Dim CurrentGrh As Grh
Dim SourceHdc As Long


'Check to make sure it is legal
If Grh.GrhIndex < 1 Then
    Exit Sub
End If
If GrhData(Grh.GrhIndex).NumFrames < 1 Then
    Exit Sub
End If

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                End If
            End If
        End If
    End If
End If

'Figure out what frame to draw (always 1 if not animated)
CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If Center Then
    If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * TilePixelWidth / 2) + TilePixelWidth / 2
    End If
    If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
    End If
End If

SourceHdc = SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum).GetDC

retcode = BitBlt(DestHdc, X, Y, GrhData(CurrentGrh.GrhIndex).pixelWidth, GrhData(CurrentGrh.GrhIndex).pixelHeight, SourceHdc, GrhData(CurrentGrh.GrhIndex).sX, GrhData(CurrentGrh.GrhIndex).sY, ROP)

SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum).ReleaseDC SourceHdc

End Sub

Sub PlayWaveDS(File As String)

    'Cylce through avaiable sound buffers
    LastSoundBufferUsed = LastSoundBufferUsed + 1
    If LastSoundBufferUsed > NumSoundBuffers Then
        LastSoundBufferUsed = 1
    End If
    
    If LoadWavetoDSBuffer(DirectSound, DSBuffers(LastSoundBufferUsed), File) Then
        DSBuffers(LastSoundBufferUsed).Play DSBPLAY_DEFAULT
    End If

End Sub

Sub PlayWaveAPI(File As String)
'*****************************************************************
'Plays a Wave using windows APIs
'*****************************************************************
Dim rc As Integer

rc = sndPlaySound(File, SND_ASYNC)

End Sub
Sub RenderScreen(TileX As Integer, TileY As Integer, PixelOffsetX As Integer, PixelOffsetY As Integer)
'***********************************************
'Draw current visible to scratch area based on TileX and TileY
'***********************************************
Dim Y As Integer    'Keeps track of where on map we are
Dim X As Integer
Dim screenminY As Integer 'Start Y pos on current screen
Dim screenmaxY As Integer 'End Y pos on current screen
Dim screenminX As Integer 'Start X pos on current screen
Dim screenmaxX As Integer 'End X pos on current screen
Dim minY As Integer 'Start Y pos on current screen + tilebuffer
Dim maxY As Integer 'End Y pos on current screen
Dim minX As Integer 'Start X pos on current screen
Dim maxX As Integer 'End X pos on current screen
Dim ScreenX As Integer 'Keeps track of where to place tile on screen
Dim ScreenY As Integer
Dim PixelOffsetXTemp As Integer 'For centering grhs
Dim PixelOffsetYTemp As Integer
Dim Moved As Byte
Dim Grh As Grh 'Temp Grh for show tile and blocked
Dim TempChar As Char

'Figure out Ends and Starts of screen
screenminY = (TileY - (WindowTileHeight \ 2))
screenmaxY = (TileY + (WindowTileHeight \ 2))
screenminX = (TileX - (WindowTileWidth \ 2))
screenmaxX = (TileX + (WindowTileWidth \ 2))

minY = screenminY - TileBufferSize
maxY = screenmaxY + TileBufferSize
minX = screenminX - TileBufferSize
maxX = screenmaxX + TileBufferSize

'Draw floor layer
ScreenY = 0
For Y = screenminY - 1 To screenmaxY + 1
    ScreenX = 0
    For X = screenminX - 1 To screenmaxX + 1
        
        'Check to see if in bounds
        If InMapBounds(X, Y) Then
    
            'Layer 1 **********************************
            
            PixelOffsetXTemp = PixelPosX(ScreenX) + PixelOffsetX + ((TileBufferSize - 1) * TilePixelWidth)
            PixelOffsetYTemp = PixelPosY(ScreenY) + PixelOffsetY + ((TileBufferSize - 1) * TilePixelHeight)
            
            'Draw
            Call DDrawGrhtoSurface(BackBufferSurface, MapData(X, Y).Graphic(1), PixelOffsetXTemp, PixelOffsetYTemp, 0, 1)
            '**********************************
            
        End If
    
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
Next Y

'Draw floor layer 2
ScreenY = 0
For Y = minY To maxY
    ScreenX = 0
    For X = minX To maxX

        'Check to see if in bounds
        If InMapBounds(X, Y) Then

            'Layer 2 **********************************
            If MapData(X, Y).Graphic(2).GrhIndex > 0 Then
            
                PixelOffsetXTemp = PixelPosX(ScreenX) + PixelOffsetX
                PixelOffsetYTemp = PixelPosY(ScreenY) + PixelOffsetY
            
                'Draw
                Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).Graphic(2), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                
            End If
            '**********************************
        End If
    
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
Next Y


'Draw transparent layers
ScreenY = 0
For Y = minY To maxY
    ScreenX = 0
    For X = minX To maxX

        'Check to see if in bounds
        If InMapBounds(X, Y) Then

            'Object Layer **********************************
            If MapData(X, Y).ObjGrh.GrhIndex > 0 Then
            
                PixelOffsetXTemp = PixelPosX(ScreenX) + PixelOffsetX
                PixelOffsetYTemp = PixelPosY(ScreenY) + PixelOffsetY
            
                'Draw
                Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                
            End If
            '**********************************
            
            
             'Char layer **********************************
            If MapData(X, Y).CharIndex > 0 Then
            
                TempChar = CharList(MapData(X, Y).CharIndex)
            
                PixelOffsetXTemp = PixelOffsetX
                PixelOffsetYTemp = PixelOffsetY
                
                Moved = 0
                'If needed, move left and right
                If TempChar.MoveOffset.X <> 0 Then
                        TempChar.Body.Walk(TempChar.Heading).Started = 1
                        PixelOffsetXTemp = PixelOffsetXTemp + TempChar.MoveOffset.X
                        TempChar.MoveOffset.X = TempChar.MoveOffset.X - (ScrollPixelsPerFrameX * Sgn(TempChar.MoveOffset.X))
                        Moved = 1
                End If
          
                'If needed, move up and down
                If TempChar.MoveOffset.Y <> 0 Then
                        TempChar.Body.Walk(TempChar.Heading).Started = 1
                        PixelOffsetYTemp = PixelOffsetYTemp + TempChar.MoveOffset.Y
                        TempChar.MoveOffset.Y = TempChar.MoveOffset.Y - (ScrollPixelsPerFrameY * Sgn(TempChar.MoveOffset.Y))
                        Moved = 1
                End If
                
                'If done moving stop animation
                If Moved = 0 And TempChar.Moving = 1 Then
                    TempChar.Moving = 0
                    TempChar.Body.Walk(TempChar.Heading).FrameCounter = 1
                    TempChar.Body.Walk(TempChar.Heading).Started = 0
                End If
                
                'Draw Body
                Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Body.Walk(TempChar.Heading), (PixelPosX(ScreenX) + PixelOffsetXTemp), PixelPosY(ScreenY) + PixelOffsetYTemp, 1, 1)
                'Draw Head
                Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Head.Head(TempChar.Heading), (PixelPosX(ScreenX) + PixelOffsetXTemp) + TempChar.Body.HeadOffset.X, PixelPosY(ScreenY) + PixelOffsetYTemp + TempChar.Body.HeadOffset.Y, 1, 0)
                
                'Refresh charlist
                CharList(MapData(X, Y).CharIndex) = TempChar
                
            End If
            '**********************************
            
            
            'Layer 3 **********************************
            If MapData(X, Y).Graphic(3).GrhIndex > 0 Then
            
                PixelOffsetXTemp = PixelPosX(ScreenX) + PixelOffsetX
                PixelOffsetYTemp = PixelPosY(ScreenY) + PixelOffsetY
            
                'Draw
                Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
            
            End If
            '**********************************
            
        End If
    
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
Next Y


'Draw blocked tiles and grid
ScreenY = 0
For Y = minY To maxY
    ScreenX = 0
    For X = minX To maxX
            
        'Check to see if in bounds
        If InMapBounds(X, Y) Then
                                
            PixelOffsetXTemp = PixelPosX(ScreenX) + PixelOffsetX
            PixelOffsetYTemp = PixelPosY(ScreenY) + PixelOffsetY
                                
            'Layer 4 **********************************
            If MapData(X, Y).Graphic(4).GrhIndex > 0 Then
                        
                'Draw
                Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).Graphic(4), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                
            End If
            '**********************************
                                
            'Draw exit
            If MapData(X, Y).TileExit.Map > 0 Then
                Grh.GrhIndex = 1
                Grh.FrameCounter = 1
                Grh.Started = 0
                Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, PixelOffsetXTemp, PixelOffsetYTemp, 0, 0)
            End If
                
            'Draw grid
            If DrawGrid = True Then
                Grh.GrhIndex = 2
                Grh.FrameCounter = 1
                Grh.Started = 0
                Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, PixelOffsetXTemp, PixelOffsetYTemp, 0, 0)
            End If

            'Show blocked tiles
            If DrawBlock = True Then
                If LegalPos(X, Y) = False Then
                    Grh.GrhIndex = 4
                    Grh.FrameCounter = 1
                    Grh.Started = 0
                    Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, PixelOffsetXTemp, PixelOffsetYTemp, 0, 0)
                End If
            End If

        End If
    
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
Next Y

End Sub

Function PixelPosX(X As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'*****************************************************************

PixelPosX = (TilePixelWidth * X) - TilePixelWidth

End Function

Function PixelPosY(Y As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'*****************************************************************

PixelPosY = (TilePixelHeight * Y) - TilePixelHeight

End Function
Sub LoadGraphics()
'*****************************************************************
'Loads all the sprites and tiles from the gif or bmp files
'*****************************************************************
Dim LoopC As Integer
Dim SurfaceDesc As DDSURFACEDESC2
Dim ddck As DDCOLORKEY
Dim ddsd As DDSURFACEDESC2

NumGrhFiles = Val(GetVar(IniPath & "Grh.ini", "INIT", "NumGrhFiles"))
ReDim SurfaceDB(1 To NumGrhFiles)



'Load the GRHx.bmps into memory
For LoopC = 1 To NumGrhFiles

    If FileExist(App.Path & GrhPath & "Grh" & LoopC & ".bmp", vbNormal) Then
        
        With ddsd
            .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        End With
        
        GetBitmapDimensions App.Path & GrhPath & "Grh" & LoopC & ".bmp", ddsd.lWidth, ddsd.lHeight
        
        Set SurfaceDB(LoopC) = DirectDraw.CreateSurfaceFromFile(App.Path & GrhPath & "Grh" & LoopC & ".bmp", ddsd)
        'Set color key
        ddck.low = 0
        ddck.high = 0
        SurfaceDB(LoopC).SetColorKey DDCKEY_SRCBLT, ddck
    End If
 
Next LoopC

End Sub
Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer) As Boolean
'*****************************************************************
'InitEngine
'*****************************************************************

Dim SurfaceDesc As DDSURFACEDESC2
Dim ddck As DDCOLORKEY

IniPath = App.Path & "\"

'Fill startup variables

DisplayFormhWnd = setDisplayFormhWnd
MainViewTop = setMainViewTop
MainViewLeft = setMainViewLeft
TilePixelWidth = setTilePixelWidth
TilePixelHeight = setTilePixelHeight
WindowTileHeight = setWindowTileHeight
WindowTileWidth = setWindowTileWidth
TileBufferSize = setTileBufferSize

'Setup borders
MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

MainViewWidth = TilePixelWidth * WindowTileWidth
MainViewHeight = TilePixelHeight * WindowTileHeight

'Resize mapdata array
ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

'Set intial user position
UserPos.X = MinXBorder
UserPos.Y = MinYBorder

'Set scroll pixels per frame
ScrollPixelsPerFrameX = 8
ScrollPixelsPerFrameY = 8

'****** INIT DirectDraw ******
' Create the root DirectDraw object
Set DirectDraw = DirectX.DirectDrawCreate("")
DirectDraw.SetCooperativeLevel DisplayFormhWnd, DDSCL_NORMAL

'Primary Surface
' Fill the surface description structure
With SurfaceDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
End With
' Create the surface
Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)

'Create Primary Clipper
Set PrimaryClipper = DirectDraw.CreateClipper(0)
PrimaryClipper.SetHWnd frmMain.hWnd
PrimarySurface.SetClipper PrimaryClipper

'Back Buffer Surface
With BackBufferRect
    .Left = 0
    .Top = 0
    .Right = TilePixelWidth * (WindowTileWidth + (2 * TileBufferSize))
    .Bottom = TilePixelHeight * (WindowTileHeight + (2 * TileBufferSize))
End With
With SurfaceDesc
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    .lHeight = BackBufferRect.Bottom
    .lWidth = BackBufferRect.Right
End With

' Create surface
Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)

'Set color key
ddck.low = 0
ddck.high = 0
BackBufferSurface.SetColorKey DDCKEY_SRCBLT, ddck

'Load graphic data into memory
Call LoadGrhData
Call LoadBodyData
Call LoadHeadData
Call LoadMapData
Call LoadGraphics

'Wave Sound
Set DirectSound = DirectX.DirectSoundCreate("")
DirectSound.SetCooperativeLevel DisplayFormhWnd, DSSCL_PRIORITY
LastSoundBufferUsed = 1

InitTileEngine = True
EngineRun = True

End Function

Sub LoadMapData()
'*****************************************************************
'Load Map.dat
'*****************************************************************

'Get Number of Maps
NumMaps = Val(GetVar(IniPath & "Map.dat", "INIT", "NumMaps"))
MapPath = GetVar(IniPath & "Map.dat", "INIT", "MapPath")

End Sub
Sub ShowNextFrame(DisplayFormTop As Integer, DisplayFormLeft As Integer)
'***********************************************
'Updates and draws next frame to screen
'***********************************************
    Static OffsetCounterX As Integer
    Static OffsetCounterY As Integer

    '****** Set main view rectangle ******
    With MainViewRect
        .Left = (DisplayFormLeft / Screen.TwipsPerPixelX) + MainViewLeft
        .Top = (DisplayFormTop / Screen.TwipsPerPixelY) + MainViewTop
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With

    '***** Check if engine is allowed to run ******
    If EngineRun Then
        
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = (OffsetCounterX - (ScrollPixelsPerFrameX * Sgn(AddtoUserPos.X)))
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = 0
                End If
            End If

            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - (ScrollPixelsPerFrameY * Sgn(AddtoUserPos.Y))
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = 0
                End If
            End If

            '****** Update screen ******
            Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
            DrawBackBufferSurface
            FramesPerSecCounter = FramesPerSecCounter + 1

    End If
    
    'Slow down engine if need be
    EvenFrameSpeed FramesPerSec, 30
        
End Sub

