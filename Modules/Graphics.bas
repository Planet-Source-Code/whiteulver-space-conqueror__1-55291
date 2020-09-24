Attribute VB_Name = "Graphics"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
  ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
  ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
  ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Type HISCORE
    hName As String * 15
    hScore As Long
    hTotalHits As Long
    hHits As Long
    hLevel As Long
    hDATE As Date
End Type

'SOUND
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
'SOUND

Public Const MAXPNAMELEN = 32

Public Type JOYCAPS
        wMid As Integer
        wPid As Integer
        szPname As String * MAXPNAMELEN
        wXmin As Integer
        wXmax As Integer
        wYmin As Integer
        wYmax As Integer
        wZmin As Integer
        wZmax As Integer
        wNumButtons As Integer
        wPeriodMin As Integer
        wPeriodMax As Integer
End Type

Public Type JOYINFO
        wXpos As Long
        wYpos As Long
        wZpos As Long
        wButtons As Long
End Type

Public Const JOY_BUTTON1 = &H1
Public Const JOY_BUTTON2 = &H2
Public Const JOY_BUTTON3 = &H4
Public Const JOY_BUTTON4 = &H8
Public Const JOYERR_BASE = 160
Public Const JOYERR_UNPLUGGED = (JOYERR_BASE + 7)
Public Const JOYSTICKID1 = 0
Public Const JOYSTICKID2 = 1


Public Declare Function joyGetDevCaps Lib "winmm.dll" Alias "joyGetDevCapsA" (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long
Public Declare Function joyGetNumDevs Lib "winmm.dll" () As Long
Public Declare Function joySetCapture Lib "winmm.dll" (ByVal hwnd As Long, ByVal uID As Long, ByVal uPeriod As Long, ByVal bChanged As Long) As Long
Public Declare Function joyReleaseCapture Lib "winmm.dll" (ByVal id As Long) As Long
Public Declare Function joyGetPos Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFO) As Long



Public mJoy As JOYINFO


Public Const HERO As Integer = 0
Public Const SMOKE As Integer = 1
Public Const ENEMY As Integer = 2
Public Const BULLET As Integer = 3
Public Const PANEL As Integer = 4



Public Const SND_ASYNC = &H1 '//lets you play a new wav sound, interrupting another
Public Const SND_LOOP = &H8     '//loops the wav sound
Public Const SND_NODEFAULT = &H2     '//if wav file not there, then make sure NOTHING plays
Public Const SND_SYNC = &H0     '//no control to program til wav is done playing
Public Const SND_NOSTOP = &H10     '//if a wav file is already playing then it wont interrupt
    
Public intL  As Boolean
Public mSprites As New mSprites
Public fireMode As Boolean
Public Dup As Boolean

'Propreties
Public DebugM As Boolean
Public SupJoystic As Long
Public eMusic As Long
Public skipFrames As Long
Public mDist As Long
Public CRAFT As Long
Public PlayerName As String * 15
Public StarF As Long
Public dLEVEL As Long
'Propreties
Public gameRun As Boolean
Public IntFrames As Long
Public backBuffer As Long
Public fCancel As Boolean
Public fEnter As Boolean
Public sSHOOT As New Wave
Public sBOOM As New Wave
Public gPause As Boolean
Public fUnload As Boolean
Public LoadedSprites As Boolean
Public mScores As HISCORE
Public cDelay As Long

Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   Key$ = LCase$(Key$)
   GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function

Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub

Public Sub playSong(sngFile As String)
mciSendString "close curSong", 0, 0, 0
strFile = Chr$(34) + sngFile + Chr$(34)
mciSendString "open " & strFile & " alias curSong", 0, 0, 0
mciSendString "Play curSong", 0, 0, 0
End Sub
Public Sub StopSong()
mciSendString "Stop curSong", 0, 0, 0
End Sub
Public Function Volume(Value As Long) As Long
mciSendString "setaudio curSong volume to " & Value, 0, 0, 0
End Function
Public Function sngStatus() As String
Dim Status As String * 128
    mciSendString "status curSong mode", Status, 128, 0
    sngStatus = Status
End Function
