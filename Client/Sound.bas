Attribute VB_Name = "Sound"
Option Explicit

Public Const NumSoundChannels = 7

Public Type MMCKINFO
    ckid As Long
    ckSize As Long
    fccType As Long
    dwDataOffset As Long
    dwFlags As Long
End Type

Public Type WAVEFORMAT
        wFormatTag As Integer
        nChannels As Integer
        nSamplesPerSec As Long
        nAvgBytesPerSec As Long
        nBlockAlign As Integer
End Type


Public DirectSound As DirectSound
Public DSBuffer(0 To NumSoundChannels) As DirectSoundBuffer

' flag values for uFlags parameter
Public Const SND_SYNC = &H0 ' play synchronously (default)
Public Const SND_ASYNC = &H1 ' play asynchronously

Public Const SND_NODEFAULT = &H2 ' silence not default, if sound not found

Public Const SND_LOOP = &H8 ' loop the sound until next sndPlaySound
Public Const SND_NOSTOP = &H10 ' don't stop any currently playing sound

Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uRetrunLength As Long, ByVal hwndCallback As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Sub CreateDSBFromWaveFile(ds As DirectSound, ByVal File As String, dsb As DirectSoundBuffer)
'*****************************************************************
'Converts a wave file to a DSB format
'*****************************************************************
Dim hFile As Long
Dim Parent As MMCKINFO ' Parent chunk wave
Dim Detail As MMCKINFO ' Detail chunks
Dim WaveF As WAVEFORMAT
Dim ptr1 As Long, ptr2 As Long, Size1 As Long, Size2 As Long
Dim BufferDesc As DSBUFFERDESC

hFile = mmioOpen(File, ByVal 0&, MMIO_READ)
 
'To point on a specific chunk, we just have to specifiy the "name" of the searched chunk and to call mmioDescend to "descend" into this chunk (chunks are structured).
Parent.fccType = mmioStringToFOURCC("WAVE", 0)
mmioDescend hFile, Parent, ByVal 0&, MMIO_FINDRIFF
 
'We are searching here for our main WAVE chunk. We can now search in this chunk, the chunks that are describing the format and that contains the wave data.
Detail.ckid = mmioStringToFOURCC("fmt ", 0)
mmioDescend hFile, Detail, Parent, MMIO_FINDCHUNK
 
'This time we searched for the "fmt " chunk that contains a WAVEFORMATEX structure that describes our wave file. We're saving now this information as this is needed to initialize our DirectSound buffer.
mmioRead hFile, WaveF, Detail.ckSize
 
'As we're done with this chunk we're going back to the parent chunk to find out the data chunk that contains wave data for the sound.
mmioAscend hFile, Detail, 0
Detail.ckid = mmioStringToFOURCC("data", 0)
mmioDescend hFile, Detail, Parent, MMIO_FINDCHUNK
 
'We are now pointing on the chunk that contains data. Before reading these data we are going to create our DirectSound buffer to receive those data.
With BufferDesc
    .dwSize = Len(BufferDesc)
    .dwFlags = DSBCAPS_CTRLDEFAULT
    ' The size of the buffer is just dtaken from the "data" chunk
    .dwBufferBytes = Detail.ckSize
    ' Point to the "fmt " chunk we read previously
    .lpwfxFormat = VarPtr(WaveF)
End With
    
ds.CreateSoundBuffer BufferDesc, dsb, Nothing
 
'Now that the DirectSoundBuffer is created, the only thing left is to read the data for the chnuk into this buffer.
dsb.Lock 0&, BufferDesc.dwBufferBytes, ptr1, Size1, ptr2, Size2, 0&
mmioRead hFile, ByVal ptr1, Size1
dsb.Unlock ptr1, Size1, ptr2, Size2
mmioClose hFile, 0&

End Sub
Sub PlayWaveDS(File As String)
'*****************************************************************
'Plays a wave using DirectSound
'*****************************************************************
Dim lngFlag As Long
Dim lngStatus As Long
Dim LoopC As Integer

'Reset any channels that are done
For LoopC = 0 To NumSoundChannels
    If Not (DSBuffer(LoopC) Is Nothing) Then
        DSBuffer(LoopC).GetStatus lngStatus
        If (lngStatus And DSBSTATUS_PLAYING) = 0 Then
            Set DSBuffer(LoopC) = Nothing
        End If
    End If
Next LoopC

'Look for open channel and play
For LoopC = 0 To NumSoundChannels
    If DSBuffer(LoopC) Is Nothing Then
        CreateDSBFromWaveFile DirectSound, File, DSBuffer(LoopC)
        DSBuffer(LoopC).Play 0, 0, lngFlag
        Exit Sub
    End If
Next LoopC

End Sub
Sub PlayMidi(File As String)
'*****************************************************************
'Plays a Midi using the MCIControl
'*****************************************************************
Dim rc As Integer

frmMain.MidiPlayer.Command = "Close"

frmMain.MidiPlayer.FileName = File
    
frmMain.MidiPlayer.Command = "Open"

frmMain.MidiPlayer.Command = "Play"

End Sub

Sub PlayWaveAPI(File As String)
'*****************************************************************
'Plays a Wave using windows APIs
'*****************************************************************
Dim rc As Integer

rc = sndPlaySound(File, SND_ASYNC)

End Sub


