VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DS_Engine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'THIS MODULE IS ENTIRELY NOT MINE

'------------------------------=
'DirectSound Class for 2D Games=
'===============================
'Author: Dave Katrowski        =
'------------------------------=

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const MAX_Channels = 8
Private Const Max_EchoSteps = 2

Private DX7 As New DirectX7, DS As DirectSound
Private BufDesc As DSBUFFERDESC, PCM As WAVEFORMATEX

Private DSB(1 To MAX_Channels) As DirectSoundBuffer
Private Echo_DSB(1 To MAX_Channels, 1 To Max_EchoSteps) As DirectSoundBuffer, dstep As Integer, sTeps As Integer, cn As Integer

Private BufferState(1 To MAX_Channels) As Boolean, st As Long
Public Sub LoadWavToChannel(ChannelNumber As Integer, Filename As String)
Set DSB(ChannelNumber) = Nothing
On Local Error GoTo BadWaveFile_Exit
Set DSB(ChannelNumber) = DS.CreateSoundBufferFromFile(Filename, BufDesc, PCM)
Let BufferState(ChannelNumber) = True: Exit Sub
BadWaveFile_Exit: Set DSB(ChannelNumber) = Nothing
End Sub
Public Sub SetFrequency(ChannelNumber As Integer, Value As Long): On Error Resume Next
If ChannelNumber > MAX_Channels Or Not BufferState(ChannelNumber) Then Exit Sub
Call DSB(ChannelNumber).SetFrequency(Value)
End Sub
Public Function GetFrequency(ChannelNumber As Integer) As Long: On Error Resume Next
If ChannelNumber > MAX_Channels Or Not BufferState(ChannelNumber) Then Exit Function
Let GetFrequency = DSB(ChannelNumber).GetFrequency
End Function
Public Sub SetVolume(ChannelNumber As Integer, Value As Long): On Error Resume Next
If ChannelNumber > MAX_Channels Or Not BufferState(ChannelNumber) Then Exit Sub
 DSB(ChannelNumber).SetVolume Value
End Sub
Public Function GetVolume(ChannelNumber As Integer) As Long
If ChannelNumber > MAX_Channels Or Not BufferState(ChannelNumber) Then Exit Function
Let GetVolume = DSB(ChannelNumber).GetVolume
End Function
Public Sub SetPan(ChannelNumber As Integer, Value As Long)
If ChannelNumber > MAX_Channels Or Not BufferState(ChannelNumber) Then Exit Sub
DSB(ChannelNumber).SetPan Value
End Sub
Public Function GetPan(ChannelNumber As Integer) As Long
If ChannelNumber > MAX_Channels Or Not BufferState(ChannelNumber) Then Exit Function
Let GetPan = DSB(ChannelNumber).GetPan
End Function
Public Sub PlaySound(ChannelNumber As Integer, Optional Looped As Boolean)
If ChannelNumber > MAX_Channels Or Not BufferState(ChannelNumber) Then Exit Sub
On Error Resume Next
If Looped Then GoTo lloop
    DSB(ChannelNumber).SetCurrentPosition 0: DSB(ChannelNumber).play DSBPLAY_DEFAULT: Exit Sub
lloop:
    DSB(ChannelNumber).SetCurrentPosition 0: DSB(ChannelNumber).play DSBPLAY_LOOPING
End Sub
Public Sub PlayEcho(ChannelNumber As Integer, step As Integer, Interval As Integer)
If ChannelNumber > MAX_Channels Or Not BufferState(ChannelNumber) Then Exit Sub
If step > Max_EchoSteps Then Let step = Max_EchoSteps
If Interval = 0 Then Let Interval = 50
For dstep = 1 To step
Set Echo_DSB(ChannelNumber, dstep) = DS.DuplicateSoundBuffer(DSB(ChannelNumber))
Echo_DSB(ChannelNumber, dstep).play DSBPLAY_DEFAULT
st = timeGetTime: Do: DoEvents: Loop While timeGetTime < st + Interval * dstep:
Next
End Sub
Public Sub StopSound(ChannelNumber As Integer)
If ChannelNumber > MAX_Channels Or Not BufferState(ChannelNumber) Then Exit Sub
DSB(ChannelNumber).Stop
End Sub
Public Sub ClearBuffer(ChannelNumber As Integer)
If ChannelNumber > MAX_Channels Or Not BufferState(ChannelNumber) Then Exit Sub
Set DSB(ChannelNumber) = Nothing
Let BufferState(ChannelNumber) = False
End Sub
Public Sub Initialize_Engine(Hwnd As Long)
Set DS = DX7.DirectSoundCreate("")
DS.SetCooperativeLevel Hwnd, DSSCL_PRIORITY

    With PCM
        .nSize = LenB(PCM)
        .nFormatTag = WAVE_FORMAT_PCM
        .nChannels = 2
        .lSamplesPerSec = 44100
        .nBitsPerSample = 16
        .nBlockAlign = .nBitsPerSample / 8 * .nChannels
        .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
    End With

BufDesc.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_STATIC

End Sub
Public Sub Terminate_Engine()

    Static I As Integer
    For I = 1 To MAX_Channels
        If BufferState(I) Then Set DSB(I) = Nothing
            For ii = 1 To Max_EchoSteps
                Set Echo_DSB(I, ii) = Nothing
            Next
    Next

Set PBUFF = Nothing
Set DS = Nothing
Set DX7 = Nothing

End Sub
Public Function GetReadableVolume(Volume As Long) As String

    Select Case Volume
        Case Is >= -10000: GetReadableVolume = "0"
        Case Is >= -3000: GetReadableVolume = "1"
        Case Is >= -2000: GetReadableVolume = "2"
        Case Is >= -1750: GetReadableVolume = "3"
        Case Is >= -1500: GetReadableVolume = "4"
        Case Is >= -1250: GetReadableVolume = "5"
        Case Is >= -1000: GetReadableVolume = "6"
        Case Is >= -750: GetReadableVolume = "7"
        Case Is >= -500: GetReadableVolume = "8"
        Case Is >= -250: GetReadableVolume = "9"
        Case 0: GetReadableVolume = "10"
    End Select

End Function

Public Function GetSpeakerConfig() As String

    Select Case DS.GetSpeakerConfig
        Case DSSPEAKER_HEADPHONE
            GetSpeakerConfig = "Speakers are set to: [Headphone]"
        Case DSSPEAKER_MONO
            GetSpeakerConfig = "Speakers are set to: [Mono]"
        Case DSSPEAKER_QUAD
            GetSpeakerConfig = "Speakers are set to: [Quadraphonic]"
        Case DSSPEAKER_SURROUND
            GetSpeakerConfig = "Speakers are set to: [Surround]"
        Case DSSPEAKER_STEREO
            GetSpeakerConfig = "Speakers are set to: [Stereo]"
     End Select
    
End Function

Public Sub SetSpeakerConfig(Command As String)
    
    Select Case LCase(Command)
        Case Is = "headphone"
            DS.SetSpeakerConfig DSSPEAKER_HEADPHONE
        Case "stereo"
            DS.SetSpeakerConfig DSSPEAKER_STEREO
        Case "mono"
            DS.SetSpeakerConfig DSSPEAKER_MONO
        Case "surround"
            DS.SetSpeakerConfig DSSPEAKER_SURROUND
        Case "quad"
            DS.SetSpeakerConfig DSSPEAKER_QUAD
        Case Else
            DS.SetSpeakerConfig DSSPEAKER_STEREO
    End Select

End Sub

Public Sub Pause(ms As Long)
Sleep ms
End Sub


