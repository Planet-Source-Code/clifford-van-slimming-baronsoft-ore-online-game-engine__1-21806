Attribute VB_Name = "Declares"
Option Explicit

'Object constants
Public Const MAX_INVENTORY_OBJS = 99
Public Const MAX_INVENTORY_SLOTS = 20

'User's inventory
Type Inventory
    OBJIndex As Integer
    Name As String
    GrhIndex As Integer
    Amount As Integer
    Equipped As Byte
End Type

'User status vars
Public UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory

Public UserName As String
Public UserPassword As String
Public UserMaxHP As Integer
Public UserMinHP As Integer
Public UserMaxMAN As Integer
Public UserMinMAN As Integer
Public UserMaxSTA As Integer
Public UserMinSTA As Integer
Public UserGLD As Long
Public UserLvl As Integer
Public UserPort As Integer
Public UserServerIP As String

'Server stuff
Public RequestPosTimer As Integer 'Used in main loop
Public stxtbuffer As String 'Holds temp raw data from server
Public SendNewChar As Boolean 'Used during login
Public Connected As Boolean 'True when connected to server
Public DownloadingMap As Boolean 'Currently downloading a map from server

'String contants
Public ENDC As String 'Endline character for talking with server
Public ENDL As String 'Holds the Endline character for textboxes

'Control
Public prgRun As Boolean 'When true the program ends

'Music stuff
Public CurMidi As String 'Keeps current MIDI file
Public LoopMidi As Byte 'If 1 current MIDI is looped


'********** OUTSIDE FUNCTIONS ***********

'For Get and Write Var
Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
