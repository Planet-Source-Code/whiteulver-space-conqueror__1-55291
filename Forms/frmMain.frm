VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Space Conqueror"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   403
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   643
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9720
      Top             =   4800
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   6060
      Left            =   0
      MouseIcon       =   "frmMain.frx":1582
      MousePointer    =   99  'Custom
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   0
      Width           =   9660
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim t As Integer, I As Integer
Dim oldX As Single, OldY As Single, curObj As Long
Dim T1 As Long, T2 As Long
Dim blT1 As Long, blT2 As Long
Dim scroll As Boolean
Const scrC  As Integer = 5
Dim tt As Integer, yy As Integer
Dim picX As Single, picY As Single

Dim frames As Long
Dim RetVal As Long
Dim numJoy As Long
Dim hScreen As Long
Dim ShowHI As Boolean
Dim hScroll As Long
Dim tScore() As HISCORE
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub JoypadCheck()
numJoy = joyGetNumDevs
If numJoy = JOYERR_UNPLUGGED Then
MsgBox "Your PC does not support Joystick or Joypad", vbInformation, App.Title
SupJoystic = 0
End If

End Sub



Private Sub Form_Activate()
Intro
End Sub

Private Sub Form_Load()
LoadGameIni
Timer1.Enabled = True
mSprites.Init App.Path & "\graphs\title.bmp", 0
mSprites.DuplicateObjT
End Sub

Private Sub LoadGameIni()
If Dir(App.Path & "\game.ini") <> "" Then

SupJoystic = GetFromINI("Settings", "Control", App.Path & "\game.ini")
If SupJoystic Then JoypadCheck

DebugM = GetFromINI("Settings", "Debug", App.Path & "\game.ini")

eMusic = GetFromINI("Settings", "Music", App.Path & "\game.ini")

skipFrames = GetFromINI("Settings", "SKIPFRAMES", App.Path & "\game.ini")

mDist = GetFromINI("Settings", "CRAFT_SPEED", App.Path & "\game.ini")

CRAFT = GetFromINI("Settings", "CRAFT", App.Path & "\game.ini")

PlayerName = Trim(GetFromINI("Settings", "PLAYER_NAME", App.Path & "\game.ini"))

StarF = GetFromINI("Settings", "STARFIELD", App.Path & "\game.ini")

dLEVEL = GetFromINI("Settings", "DIFFICULTY", App.Path & "\game.ini")

End If

End Sub
Private Sub SaveGameIni()

WriteToINI "Settings", "Control", CStr(SupJoystic), App.Path & "\game.ini"

WriteToINI "Settings", "Control", CStr(eMusic), App.Path & "\game.ini"

End Sub
Private Sub LoadSprites()
mSprites.createFONT Pic1.hdc, 1



mSprites.InitMulti App.Path & "\graphs\craft" & CRAFT & ".spr", 0
mSprites.DuplicateObj HERO, True
mSprites.setCoords 100, 223, HERO, "HERO", HERO


mSprites.InitMulti App.Path & "\graphs\smoke.spr", 0
mSprites.DuplicateObjS 0, True
mSprites.setCoordsS 0, 0, 0


sSHOOT.InitSound App.Path & "\sounds\shoot.wav", Me
sBOOM.InitSound App.Path & "\sounds\boom.wav", Me

mSprites.Init App.Path & "\graphs\ENEMYA.BMP", 0
mSprites.DuplicateObjH 0, True
mSprites.setCoordsH -550, 250, 0, "ENEMY"

mSprites.Init App.Path & "\graphs\FIRE1.BMP", 0
mSprites.DuplicateObjB 0, True

'mSprites.mCreate Pic1.hdc, 250, 150, "PANEL"

If skipFrames = 1 Then
    cDelay = 0
Else
    cDelay = 30
End If
'Pic1.Refresh
LoadedSprites = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
gPause = False
fCancel = True
fUnload = True
Timer1.Enabled = False
StopSong
'SaveGameIni
End Sub

Private Sub PrepareGame()
backBuffer = mSprites.LoadGraphicDC(App.Path & "\graphs\WORLD.BMP")
If LoadedSprites = False Then LoadSprites
fCancel = False
End Sub

Private Sub mainLoop()

        T2 = GetTickCount
gameRun = True
        Do
        DoEvents
        If gPause Then GoTo SkipGame:
        
        T1 = GetTickCount

        If T1 - T2 > cDelay Then '1
        
            If fCancel Then Exit Sub
            
            Pic1.Cls


'-----------------------------------------------------------

If StarF = 0 Then
BitBlt Pic1.hdc, tt, yy, 640, 400, backBuffer, 0, 0, vbSrcCopy 'Background
Else
mSprites.Starfield Pic1.hdc
End If
mSprites.click_event Pic1.hdc
mSprites.CollisionDetection
mSprites.reDraw Pic1.hdc
            
'-----------------------------------------------------------
  
            T2 = GetTickCount
            
            'Pic1.Refresh
            frames = frames + 1
    End If '1
SkipGame:
        Loop Until fCancel = True
        BitBlt backBuffer, 0, 0, 640, 400, 0, 0, 0, vbBlack
Intro
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If LoadedSprites Then
        Set mSprites = Nothing
    End If
sSHOOT.Terminate
End Sub




Private Sub Pic1_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

Case Is = vbKeyReturn

fEnter = True

Case Is = vbKeyDown

If ShowHI Then
hScroll = hScroll + 24
If hScroll >= 640 Then hScroll = -400
End If

Case Is = vbKeyUp

If ShowHI Then
hScroll = hScroll - 24
If hScroll < 8 Then hScroll = 0
End If

Case Is = vbKeyH

Dim h As Integer
Dim fLen As Long

If ShowHI = False Then
    ShowHI = True
Else
   ShowHI = False
End If


fLen = FileLen(App.Path & "\hiscores.dat")
If fLen <> 0 Then

'h = 1

Open App.Path & "\hiscores.dat" For Binary Access Read As #1
mSprites.LoadScreen hScreen, ((fLen / Len(mScores)) + 1) * 24
ReDim tScore((fLen / Len(mScores)))
Do
Get #1, , mScores
    tScore(h) = mScores
    h = h + 1
Loop Until h > (fLen / Len(mScores))
SortArray
Close #1

End If

Case Is = vbKeyP
If gameRun Then
If gPause = False Then
    gPause = True
    mSprites.mDrawText Pic1.hdc, 260, 200, "PAUSE", 20, RGB(255, 0, 0)
    Pic1.Refresh
    Timer1.Enabled = False
Else
    gPause = False
    Timer1.Enabled = True
End If
End If
End Select

End Sub

Private Sub Pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picX = X: picY = Y
End Sub

Private Sub Timer1_Timer()
If skipFrames = 0 Then
If frames > 35 Then cDelay = cDelay + 10
If frames <= 30 Then cDelay = cDelay - 10
If cDelay < 0 Then cDelay = 0
End If
IntFrames = frames

frames = 0
If eMusic = 1 And gameRun Then
If Mid(sngStatus, 1, 7) = "stopped" Then
playSong App.Path & "\sounds\game.mp3"
End If
End If
End Sub


Private Sub Intro()
Dim blinkTime As Long
Static zC As Long, tBlink As Long

zC = 640
T2 = GetTickCount
blT2 = GetTickCount
Do

DoEvents

T1 = GetTickCount
blT1 = GetTickCount

If T1 - T2 > cDelay Then '1
    
    
    If fEnter Then Exit Do
    If fUnload Then Exit Sub
    
    
    Pic1.Cls
    
    
    If ShowHI Then 'HI
        BitBlt Pic1.hdc, 0, -hScroll, 640, 1200, hScreen, 0, 0, vbSrcCopy
        mSprites.SolidRect Pic1.hdc, 0, 0, 640, 24
        'Pic1.Line (0, 0)-(640, 24), RGB(0, 0, 0), BF
        mSprites.mDrawText Pic1.hdc, 50, 0, "PLAYER NAME      HISCORE    LEVEL  HITS        DATE", 20, RGB(255, 0, 0)
    
    Else
        If zC <= 0 Then zC = 0
        mSprites.Starfield Pic1.hdc
        mSprites.sDraw Pic1.hdc, zC
        
        If CInt(blinkTime / 1000) = tBlink And blinkTime <> 0 Then
            If skipFrames = 1 Or IntFrames < 40 Then
                mSprites.mDrawText Pic1.hdc, 250, 250, "Press ENTER", 20, RGB(255, 0, 0)
            Else
                mSprites.mDrawText Pic1.hdc, 250, 250, "PLEASE WAIT...", 20, RGB(255, 0, 0)
            End If
        End If
        

    End If 'HI

mSprites.mDrawText Pic1.hdc, 10, 380, "Programming By Whiteulver ©2004", 10, RGB(255, 0, 0)

T2 = GetTickCount
    
    
    frames = frames + 1
    zC = zC - 15

End If '1

If (blT1 - blT2) >= 3000 Then
    blinkTime = blT1 - blT2
    tBlink = blinkTime \ 1000
End If

Loop Until fEnter = True

'
Pic1.Cls
BitBlt Pic1.hdc, 0, 0, 640, 400, backBuffer, 0, 0, vbSrcCopy
mSprites.mDrawText Pic1.hdc, 250, 250, "LOADING... ", 20, RGB(255, 0, 0)
Pic1.Refresh
If eMusic = 1 Then
    Volume 500
    DoEvents
    playSong App.Path & "\sounds\game.mp3"
End If
PrepareGame
mSprites.GameCondition 0, dLEVEL, True

mainLoop
'
End Sub

Private Sub SortArray()

Dim min As Long, max As Long
Dim I As Long
Dim t As Long
Dim k As Long
Dim dp As Long
Dim sm() As HISCORE
Dim h As Long
Dim cCol As Long
ReDim sm(UBound(tScore))

'find max
t = 0
For I = 0 To UBound(tScore)
If tScore(I).hScore >= t Then max = tScore(I).hScore: t = tScore(I).hScore
Next I

'find min
t = max
For I = 0 To UBound(tScore)
If tScore(I).hScore <= t Then min = tScore(I).hScore: t = tScore(I).hScore
Next I


'sort by max
For I = 0 To UBound(tScore)
t = min
For k = 0 To UBound(tScore)
If tScore(k).hScore >= t Then sm(I) = tScore(k): dp = k: t = tScore(k).hScore
Next k
tScore(dp).hScore = min
h = h + 1

If h > UBound(tScore) Then Exit Sub

If h > 0 Then cCol = RGB(0, 255, 0)
If h > 5 Then cCol = RGB(0, 200, 0)
If h > 10 Then cCol = RGB(0, 150, 0)
If h > 15 Then cCol = RGB(0, 100, 0)
If h > 25 Then cCol = RGB(0, 50, 0)
    mSprites.mDrawText hScreen, 10, (h * 24), CStr(h), 20, cCol
    mSprites.mDrawText hScreen, 50, (h * 24), (sm(I).hName), 20, cCol
    mSprites.mDrawText hScreen, 250, (h * 24), CStr(sm(I).hScore), 20, cCol
    mSprites.mDrawText hScreen, 375, (h * 24), CStr(sm(I).hLevel), 20, cCol
    mSprites.mDrawText hScreen, 415, (h * 24), CStr(Round((sm(I).hHits * 100) / sm(I).hTotalHits, 0)) & " %", 20, cCol
    mSprites.mDrawText hScreen, 500, (h * 24), CStr(sm(I).hDATE), 20, cCol
Next I

End Sub
