VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mSprites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type xPOINTAPI
        X As Long
        Y As Long
        clr As Long
End Type

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type tFrame '14 bytes
    mBank As Integer
    sndBank As Integer
    mX As Integer
    mY As Integer
    velX As Integer
    velY As Integer
    tick As Long
    hit As Byte
    Def As Byte
End Type

Private Type tAction
    tActID As Byte
    numOS As Byte
    tLenght As Long
    actFrm() As tFrame
End Type

Private Type ObjMulti
    oName As String * 12
    Class As Integer
    NOPMu As Integer
'
    hdc() As Long
    hDC_H() As Long
    hDC_M() As Long
'
    oX As Long
    oY As Long
    bX() As Long
    bY() As Long
    mW() As Long
    mH() As Long
    F As Integer
    lPath() As POINTAPI
    oRECT() As RECT
'
    animR As Boolean
    IHover As Boolean
    LHover As Boolean
End Type

Private Type Obj
    oName As String * 12
    NOP As Long
    
    hdc As Long
    hDC_H As Long
    hDC_M As Long
    bCnt As Long

    IHover As Boolean
    animR As Boolean
    
    mW As Long
    mH As Long
    mPoints() As POINTAPI
    oX As Long
    oY As Long
    oRECT As RECT
    lPath() As POINTAPI
End Type



Private Type Ar
    aN() As POINTAPI
End Type

Private Type KeyboardInput
    btnDown As Boolean
    btnRight As Boolean
    btnUp As Boolean
    btnLeft As Boolean
    btnA As Boolean
    btnB As Boolean
    btnC As Boolean
    btnD As Boolean
    btnE As Boolean
    btnF As Boolean
    btnG As Boolean
    btnH As Boolean
    btnI As Boolean
    btnJ As Boolean
    btnK As Boolean
    btnL As Boolean
    btnM As Boolean
    btnN As Boolean
    btnO As Boolean
    btnP As Boolean
    btnQ As Boolean
    btnR As Boolean
    btnS As Boolean
    btnT As Boolean
    btnU As Boolean
    btnV As Boolean
    btnW As Boolean
    btnX As Boolean
    btnY As Boolean
    btnZ As Boolean
    btnAlt As Boolean
    btnControl As Boolean
    btnReturn As Boolean
End Type

Dim myKey As KeyboardInput
Private Const RadToDegFuncX As Double = 57.2957795130824
Private Const DegToRadAngleX As Double = 1.74532925199433E-02

Private cAction As tAction

'lfCharSet Constants
Private Const ANSI_CHARSET = 0
Private Const DEFAULT_CHARSET = 1
Private Const SYMBOL_CHARSET = 2
Private Const SHIFTJIS_CHARSET = 128
Private Const HANGEUL_CHARSET = 129
Private Const OEM_CHARSET = 255

'lfPitchAndFamily Const
Private Const DEFAULT_PITCH = 0
Private Const FIXED_PITCH = 1
Private Const VARIABLE_PITCH = 2
Private Const FF_DECORATIVE = 80
Private Const FF_DONTCARE = 0
Private Const FF_MODERN = 48
Private Const FF_ROMAN = 16
Private Const FF_SCRIPT = 64
Private Const FF_SWISS = 32

Private Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hdc As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, ByVal lParam As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * 32
End Type

Private Type BITMAPFILEHEADER    '14 bytes
   bfType As Integer
   bfSize As Long
   bfReserved1 As Integer
   bfReserved2 As Integer
   bfOffBits As Long
End Type

Private Type BITMAPINFOHEADER   '40 bytes
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

Private Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type

Private Type BITMAPINFO_8
   bmiHeader As BITMAPINFOHEADER
   bmiColors(0 To 255) As RGBQUAD
End Type

Private Type BITMAPINFO_24
   bmiHeader As BITMAPINFOHEADER
End Type

Private Declare Function SetDIBits_8 Lib "gdi32" Alias "SetDIBits" _
  (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
  ByVal nNumScans As Long, lpBits As Any, lpbi As BITMAPINFO_8, _
  ByVal wUsage As Long) As Long

Private Declare Function SetDIBits_24 Lib "gdi32" Alias "SetDIBits" _
  (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
  ByVal nNumScans As Long, lpBits As Any, lpbi As BITMAPINFO_24, _
  ByVal wUsage As Long) As Long

Private Declare Function ArcTo Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetArcDirection Lib "gdi32" (ByVal hdc As Long, ByVal ArcDirection As Long) As Long
Private Declare Function Chord Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long

Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function PathToRegion Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
  ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
  ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
  ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
  ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long


Private Const DIB_RGB_COLORS As Long = 0

Const H_COLOR = 38400
Const N_COLOR = 13160660

Private TIMEFORENEMIES As Long


Private cDist As Integer
Private charT As Integer
Public UText As String 'Text of Sprite
Public oID As Integer 'Identify of sprite
Public curX As Long, curY As Long
Public ccA As Integer 'Multiple of Font
Public anim As Boolean 'Check if sprite contane more than one image
Public animS As Boolean
Public gCnt As Long ', bCnt As Long
Public gStop As Boolean 'Trigger to stop animation
Public oSpeed As Integer
Public n As Integer, cn As Integer, rnt As Integer 'some counters
Public absX As Long, absY As Long 'Absolute coordinetes of object(sprite)
'
Dim mRGN As Long 'region to use with structure fillrgn
Dim mRECT As RECT
Dim RetVal As Long
Dim prev As POINTAPI 'previous point for "MoveToEx" structure
Dim curFont As LOGFONT
Dim font As Long 'Handle to send the fonts
Dim mColor As Long
Dim ms1Color As Long
Dim ms2color As Long
Dim hColor As Long
Dim hs1Color As Long
Dim hs2Color As Long

Dim mObj As Obj 'Single Image Object
Dim pObj As Obj 'Panel
Dim gObj As ObjMulti 'Multi image object

Dim bObj As Obj 'Bullet Global
Dim hsObj As Obj 'Hostile Global
Dim tObj As Obj 'Title Global
Dim gCraftObj As ObjMulti 'Craft Global
Dim gSmokeObj As ObjMulti 'Smoke Global

Dim buObj(0 To 15) As Obj 'Bullets
Dim cObj() As ObjMulti 'Hero Crafts
Dim sObj() As ObjMulti 'Smokes
Dim hObj() As Obj 'Hostile Crafts

Dim emptyObjM As ObjMulti 'Null
Dim emptyObj As Obj 'Null

Dim I As Integer 'counter
Dim brush As Long, pen As Long 'Brush and pen for drawing
Dim ScreenDC As Long, picbmp As Long, picbmpH As Long, picbmpM As Long
Dim pA As Long

'
Private canWalk As Boolean
Private glbPoint() As POINTAPI
Private glbCnt As Long
Private gArray() As Ar
Private sRECT As RECT
Private tempRect As RECT

Dim bT1 As Long, bT2  As Long
Dim eT1 As Long, eT2  As Long
Dim dT1 As Long, dT2  As Long
Dim pT1 As Long, pT2  As Long
Dim fT1 As Long, fT2 As Long

Dim Hits As Long
Dim GamePause As Boolean
Dim HeroKill As Boolean
Dim gFinale As Boolean
Dim HitsPerLevel As Long, curLevel As Long
Dim craftLife As Integer, tBull As Integer
Dim curEn As Integer
Dim totalSC As Long
Dim AllEnemies As Long, Allbul As Long

Private Type Star
    X As Long
    Y As Long
    Speed As Long
End Type

'Star field array
Dim Stars(49) As Star
Const MaxSize As Long = 5
Const MaxSpeed As Long = 25

Public Sub GameCondition(cLevel As Long, dif As Long, Optional fTime As Boolean)
If fTime Then

'curLevel = 0
AllEnemies = 0
Allbul = 0
cLevel = cLevel + 1
HitsPerLevel = cLevel * 100
craftLife = 3
Hits = 0
curLevel = cLevel
gFinale = False
End If

Select Case dif
Case Is = 1

Select Case cLevel
Case Is = 1
TIMEFORENEMIES = 500
Case Is = 3
TIMEFORENEMIES = 400
Case Is = 5
TIMEFORENEMIES = 300
Case Is = 7
TIMEFORENEMIES = 200
End Select

Case Is = 2

Select Case cLevel
Case Is = 1
TIMEFORENEMIES = 400
Case Is = 3
TIMEFORENEMIES = 300
Case Is = 5
TIMEFORENEMIES = 200
Case Is = 7
TIMEFORENEMIES = 100
End Select

Case Is = 3

Select Case cLevel
Case Is = 1
TIMEFORENEMIES = 350
Case Is = 3
TIMEFORENEMIES = 200
Case Is = 5
TIMEFORENEMIES = 100
Case Is = 7
TIMEFORENEMIES = 50
End Select

End Select
End Sub


Friend Sub DuplicateObj(objN As Integer, Optional fTime As Boolean)
If fTime Then
ReDim Preserve cObj(objN)
ReDim gCraftObj.hdc(UBound(gObj.hdc))
ReDim gCraftObj.hDC_H(UBound(gObj.hDC_H))
ReDim gCraftObj.hDC_M(UBound(gObj.hDC_M))
ReDim gCraftObj.bX(UBound(gObj.bX))
ReDim gCraftObj.bY(UBound(gObj.bY))
ReDim gCraftObj.mH(UBound(gObj.mH))
ReDim gCraftObj.mW(UBound(gObj.mW))

ReDim emptyObjM.hdc(UBound(gCraftObj.hdc))
ReDim emptyObjM.hDC_H(UBound(gCraftObj.hDC_H))
ReDim emptyObjM.hDC_M(UBound(gCraftObj.hDC_M))
ReDim emptyObjM.bX(UBound(gCraftObj.bX))
ReDim emptyObjM.bY(UBound(gCraftObj.bY))
ReDim emptyObjM.mH(UBound(gCraftObj.mH))
ReDim emptyObjM.mW(UBound(gCraftObj.mW))

gCraftObj = gObj
cObj(objN) = gCraftObj


Else
    ReDim Preserve cObj(objN)
    cObj(objN) = gCraftObj

End If
End Sub



Public Sub DuplicateObjS(objN As Integer, Optional fTime As Boolean)
If fTime Then
ReDim Preserve sObj(objN)
ReDim gSmokeObj.hdc(UBound(gObj.hdc))
ReDim gSmokeObj.hDC_H(UBound(gObj.hDC_H))
ReDim gSmokeObj.hDC_M(UBound(gObj.hDC_M))
ReDim gSmokeObj.bX(UBound(gObj.bX))
ReDim gSmokeObj.bY(UBound(gObj.bY))
ReDim gSmokeObj.mH(UBound(gObj.mH))
ReDim gSmokeObj.mW(UBound(gObj.mW))

ReDim emptyObjM.hdc(UBound(gSmokeObj.hdc))
ReDim emptyObjM.hDC_H(UBound(gSmokeObj.hDC_H))
ReDim emptyObjM.hDC_M(UBound(gSmokeObj.hDC_M))
ReDim emptyObjM.bX(UBound(gSmokeObj.bX))
ReDim emptyObjM.bY(UBound(gSmokeObj.bY))
ReDim emptyObjM.mH(UBound(gSmokeObj.mH))
ReDim emptyObjM.mW(UBound(gSmokeObj.mW))

gSmokeObj = gObj
sObj(objN) = gSmokeObj

Else
    ReDim Preserve sObj(objN)
    sObj(objN) = gSmokeObj
End If
End Sub
Public Sub DuplicateObjH(curObj As Integer, Optional fTime As Boolean)
If fTime Then
    hsObj = mObj
Else
    ReDim Preserve hObj(curObj)
    hObj(curObj) = hsObj
End If
End Sub

Public Sub DuplicateObjB(curObj As Integer, Optional fTime As Boolean)
If fTime Then
    bObj = mObj
Else
'    ReDim Preserve buObj(curObj)
    buObj(curObj) = bObj
    buObj(curObj).oX = cObj(HERO).oX
    buObj(curObj).oY = cObj(HERO).oY
End If
End Sub
Public Sub DuplicateObjT()
    tObj = mObj
End Sub
Private Sub drawLine(hHandle As Long, sLineX As Long, sLineY As Long, eLineX As Long, eLineY As Long, iobj As Integer, multi As Boolean, Optional mReset As Boolean)
'good old bresenhams line drawing algorithm
Dim xC#, yC#, DX#, dy#
Dim xinc1#, xinc2#, yinc1#, yinc2#
Dim den#, num#, numadd#, numpixels#, curpixel#

ReDim Preserve gArray(iobj)

If mReset = False Then glbCnt = 0

DX = Abs(eLineX - sLineX)
dy = Abs(eLineY - sLineY)

xC = sLineX
yC = sLineY

If (eLineX >= sLineX) Then  '  // The x-values are increasing
  xinc1 = 1
  xinc2 = 1
Else                '  // The x-values are decreasing
  xinc1 = -1
  xinc2 = -1
End If

If (eLineY >= sLineY) Then  '  // The y-values are increasing
  yinc1 = 1
  yinc2 = 1
Else                '  // The y-values are decreasing
  yinc1 = -1
  yinc2 = -1
End If

If (DX >= dy) Then  '  // There is at least one x-value for every y-value
  xinc1 = 0         '  // Don't change the x when numerator >= denominator
  yinc2 = 0         '  // Don't change the y for every iteration
  den = DX
  num = DX / 2
  numadd = dy
  numpixels = DX    '  // There are more x-values than y-values
Else                '  // There is at least one y-value for every x-value
  xinc2 = 0         '  // Don't change the x for every iteration
  yinc1 = 0         '  // Don't change the y when numerator >= denominator
  den = dy
  num = dy / 2
  numadd = DX
  numpixels = dy    '  // There are more y-values than x-values
End If
For curpixel = 0 To numpixels
If DebugM Then SetPixelV hHandle, xC, yC, RGB(255, 0, 0)
glbCnt = glbCnt + 1

ReDim Preserve gArray(iobj).aN(glbCnt)

gArray(iobj).aN(glbCnt).X = xC: gArray(iobj).aN(glbCnt).Y = yC
          num = num + numadd ' // Increase the numerator by the top of the fraction
          If num >= den Then ' // Check if numerator >= denominator
            num = num - den '  // Calculate the new numerator value
            xC = xC + xinc1   '  // Change the x as appropriate
            yC = yC + yinc1 '  // Change the y as appropriate
          End If
          xC = xC + xinc2     '  // Change the x as appropriate
          yC = yC + yinc2     '  // Change the y as appropriate

Next curpixel
gArray(iobj).aN(0).X = gArray(iobj).aN(1).X
gArray(iobj).aN(0).Y = gArray(iobj).aN(1).Y

pA = UBound(gArray(iobj).aN)

If multi Then
ReDim Preserve hObj(iobj).lPath(pA)
hObj(iobj).lPath = gArray(iobj).aN
Else
ReDim Preserve buObj(iobj).lPath(pA)
buObj(iobj).lPath = gArray(iobj).aN
End If
End Sub


Friend Sub Init(ByVal filename As String, sChar As Integer)
anim = False
charT = sChar

'Loads a filename into a picture in memory.
    Dim FileNo As Integer
    Dim ScreenDC As Long
    Dim FileHead As BITMAPFILEHEADER
    Dim bmpInfo_24 As BITMAPINFO_24
    Dim bmpInfo_8 As BITMAPINFO_8
    Dim bmpInfo_8M As BITMAPINFO_8
    Dim bmpInfo_8H As BITMAPINFO_8
    Dim bmpInfoHead As BITMAPINFOHEADER
    Dim PicBytes() As Byte
    Dim RetVal As Long
    Dim PicLen As Long
    
    On Error GoTo Err_Init
    
    'Make sure the file exists
    If Dir(filename, vbArchive Or vbNormal) = "" Then
        MsgBox "You must pass a valid BMP file to this routine!"
        Exit Sub
    End If
       
    'Read in the file data
    FileNo = FreeFile
    Open filename For Binary Access Read As #FileNo
    
     DoEvents
    'Retrieve the bitmap information
    Get #FileNo, , FileHead
    Get #FileNo, 15, bmpInfoHead
    
    PicLen = FileHead.bfSize
    
    'Load the picture into the memory bitmap
    Select Case bmpInfoHead.biBitCount
        Case Is = 8
            Get #FileNo, 15, bmpInfo_8
        Case Is = 24
            Get #FileNo, 15, bmpInfo_24
    End Select
    
    bmpInfo_8M = bmpInfo_8
    
    With bmpInfo_8.bmiColors(0)
        .rgbRed = 0
        .rgbGreen = 0
        .rgbBlue = 0
    End With
    
    With bmpInfo_8M.bmiColors(0)
        .rgbRed = 255
        .rgbGreen = 255
        .rgbBlue = 255
    End With

    bmpInfo_8H = bmpInfo_8
    
    With bmpInfo_8H.bmiColors(0)
        .rgbRed = 255
        .rgbGreen = 255
        .rgbBlue = 255
    End With
    
    For I = 1 To 255
    With bmpInfo_8H.bmiColors(I)
        .rgbRed = 255
        .rgbGreen = 255
        .rgbBlue = 0
    End With
    Next
    
    ReDim PicBytes(0 To PicLen - 44)
    With mObj
    Get #FileNo, , PicBytes
    'Create the memory bitmap
    ScreenDC = GetDC(0)
    .hdc = CreateCompatibleDC(ScreenDC)
    .hDC_M = CreateCompatibleDC(ScreenDC)
    .hDC_H = CreateCompatibleDC(ScreenDC)
    .mW = bmpInfoHead.biWidth
    .mH = bmpInfoHead.biHeight
    picbmp = CreateCompatibleBitmap(ScreenDC, .mW, .mH)
    picbmpM = CreateCompatibleBitmap(ScreenDC, .mW, .mH)
    picbmpH = CreateCompatibleBitmap(ScreenDC, .mW, .mH)
    SelectObject .hdc, picbmp
    SelectObject .hDC_M, picbmpM
    SelectObject .hDC_H, picbmpH
    'Load the picture into the memory bitmap
    Select Case bmpInfoHead.biBitCount
        Case Is = 8
            RetVal = SetDIBits_8(.hdc, picbmp, 0, .mH, PicBytes(0), bmpInfo_8, DIB_RGB_COLORS)
            RetVal = SetDIBits_8(.hDC_M, picbmpM, 0, .mH, PicBytes(0), bmpInfo_8M, DIB_RGB_COLORS)
            RetVal = SetDIBits_8(.hDC_H, picbmpH, 0, .mH, PicBytes(0), bmpInfo_8H, DIB_RGB_COLORS)
        Case Is = 24
            RetVal = SetDIBits_24(.hdc, picbmp, 0, .mH, PicBytes(0), bmpInfo_24, DIB_RGB_COLORS)
            RetVal = SetDIBits_24(.hDC_H, picbmpH, 0, .mH, PicBytes(0), bmpInfo_24, DIB_RGB_COLORS)
    End Select
    
    'Clean up the bitmap
    RetVal = ReleaseDC(0, ScreenDC)
    'Close the file
    Close #FileNo
    DeleteObject picbmp
    DeleteObject picbmpM
    DeleteObject picbmpH
    
    End With
    DuplicateObjH sChar
    'BitBlt cHDC, pX, pY, mObj.mW, mObj.mH, mObj.hDC_M, 0, 0, vbSrcAnd
    'BitBlt cHDC, pX, pY, mObj.mW, mObj.mH, mObj.hdc, 0, 0, vbSrcPaint
    

    Exit Sub
    
Err_Init:
    MsgBox Err.Number & " - " & Err.Description
    Resume Next
End Sub


Friend Sub InitMulti(ByVal mFile As String, sChar As Integer)
anim = True
charT = sChar

'Loads a filename into a picture in memory.
    Dim FileNo As Integer
    Dim ScreenDC As Long
    Dim FileHead As BITMAPFILEHEADER
    Dim bmpInfo_24 As BITMAPINFO_24
    Dim bmpInfo_8 As BITMAPINFO_8
    Dim bmpInfo_8M As BITMAPINFO_8
    Dim bmpInfo_8H As BITMAPINFO_8
    Dim bmpInfoHead As BITMAPINFOHEADER
    Dim PicBytes() As Byte
    Dim RetVal As Long
    Dim PicOffset As Long, PicLen As Long
    
    n = 0
    'oSpeed = sspeed
    anim = True
    
With gObj
    
    .NOPMu = 0: I = 0
    
    'On Error GoTo Err_Init
    
'    absX = pX: absY = pY
    
    'Make sure the file exists
    If Dir(mFile, vbArchive Or vbNormal) = "" Then
        MsgBox "You must pass a valid BMP file to this routine!"
        Exit Sub
    End If
    
    Open mFile For Binary Access Read As #1
        Get #1, 1, .NOPMu
    Close #1
    ReDim .hdc(.NOPMu)
    ReDim .hDC_M(.NOPMu)
    ReDim .hDC_H(.NOPMu)
    ReDim .mW(.NOPMu)
    ReDim .mH(.NOPMu)
    ReDim .bX(.NOPMu)
    ReDim .bY(.NOPMu)
    ReDim .oRECT(.NOPMu)
    
    cmdLoadFile mFile
    
    PicOffset = 3
    
    'Read in the file data
    FileNo = FreeFile
    Open mFile For Binary Access Read As #FileNo
    Do Until PicOffset > FileLen(mFile)
   
    DoEvents
    
    'Retrieve the bitmap information
    Get #FileNo, PicOffset, FileHead
    Get #FileNo, PicOffset + 14, bmpInfoHead
    
    PicLen = FileHead.bfSize
    
    'Load the picture into the memory bitmap
    Select Case bmpInfoHead.biBitCount
        Case Is = 8
            Get #FileNo, PicOffset + 14, bmpInfo_8
        Case Is = 24
            Get #FileNo, PicOffset + 14, bmpInfo_24
    End Select
    
    bmpInfo_8M = bmpInfo_8

    'Change the backcolor of main sprite pic to black
    With bmpInfo_8.bmiColors(0)
        .rgbRed = 0
        .rgbGreen = 0
        .rgbBlue = 0
    End With
    
    'Change the backcolor of mask sprite pic to white
    With bmpInfo_8M.bmiColors(0)
        .rgbRed = 255
        .rgbGreen = 255
        .rgbBlue = 255
    End With
    
    'bmpInfo_8H = bmpInfo_8
    'With bmpInfo_8H.bmiColors(6)
    '    .rgbRed = 255
    '    .rgbGreen = 255
    '    .rgbBlue = 255
    'End With

    ReDim PicBytes(0 To PicLen - 44)
    
    Get #FileNo, , PicBytes
    
    'Create the memory bitmap
    ScreenDC = GetDC(0)
    .hdc(I) = CreateCompatibleDC(ScreenDC)
    .hDC_M(I) = CreateCompatibleDC(ScreenDC)
    .hDC_H(I) = CreateCompatibleDC(ScreenDC)
    .mW(I) = bmpInfoHead.biWidth
    .mH(I) = bmpInfoHead.biHeight
    picbmp = CreateCompatibleBitmap(ScreenDC, .mW(I), .mH(I))
    picbmpM = CreateCompatibleBitmap(ScreenDC, .mW(I), .mH(I))
    picbmpH = CreateCompatibleBitmap(ScreenDC, .mW(I), .mH(I))
    SelectObject .hdc(I), picbmp
    SelectObject .hDC_M(I), picbmpM
    SelectObject .hDC_H(I), picbmpH
    
    'Load the picture into the memory bitmap

    SetDIBits_8 .hdc(I), picbmp, 0, .mH(I), PicBytes(0), bmpInfo_8, DIB_RGB_COLORS
    SetDIBits_8 .hDC_M(I), picbmpM, 0, .mH(I), PicBytes(0), bmpInfo_8M, DIB_RGB_COLORS
    SetDIBits_8 .hDC_H(I), picbmpH, 0, .mH(I), PicBytes(0), bmpInfo_8H, DIB_RGB_COLORS
    
    
    '.bX(i) = cAction.actFrm(i).mX
    '.bY(i) = cAction.actFrm(i).mY
    
    'Clean up the bitmap
    RetVal = ReleaseDC(0, ScreenDC) 'frmMain.Pic1.hwnd
    DeleteObject picbmp
    DeleteObject picbmpM
    DeleteObject picbmpH
    
    PicOffset = PicOffset + FileHead.bfSize
    
    I = I + 1
    
    Loop
    
    'Close the file
    Close #FileNo
    .NOPMu = I - 1
    DeleteObject picbmp
    DeleteObject picbmpM
    DeleteObject picbmpH

End With

    
    Exit Sub
    
Err_Init:
    MsgBox Err.Number & " - " & Err.Description
    Resume Next

End Sub


Public Sub mCreate(cHDC As Long, pX As Long, pY As Long, H_STRING As String)
anim = False
pObj.NOP = 3
ReDim pObj.mPoints(3)

UText = H_STRING
Dim tmpString As String

mColor = RGB(212, 208, 200) '13160660
ms1Color = RGB(255, 255, 255)
ms2color = RGB(180, 180, 180)
hColor = RGB(150, 0, 0) ' 38400
hs1Color = RGB(255, 200, 200)
hs2Color = RGB(150, 0, 0)

'----coordinates----
With pObj
.oX = pX: .oY = pY
    
    .mPoints(0).X = 0: .mPoints(0).Y = 200
    .mPoints(1).X = 0: .mPoints(1).Y = 0
    .mPoints(2).X = 120: .mPoints(2).Y = 0
    .mPoints(3).X = 120: .mPoints(3).Y = 200
    .mW = (.mPoints(2).X - .mPoints(0).X)
    .mH = (.mPoints(0).Y - .mPoints(2).Y)
    
    ScreenDC = GetDC(0)
    .hdc = CreateCompatibleDC(ScreenDC)
    .hDC_H = CreateCompatibleDC(ScreenDC)
    .hDC_M = CreateCompatibleDC(ScreenDC)
    picbmp = CreateCompatibleBitmap(ScreenDC, .mW, .mH)
    picbmpH = CreateCompatibleBitmap(ScreenDC, .mW, .mH)
    picbmpM = CreateCompatibleBitmap(ScreenDC, .mW, .mH)
    SelectObject .hdc, picbmp
    SelectObject .hDC_H, picbmpH
    SelectObject .hDC_M, picbmpM

    '----fill----
    SetBkColor .hdc, RGB(0, 0, 0)
    SetBkColor .hDC_H, RGB(0, 0, 0)
    
    SetBkMode .hdc, 1 'Used for transparent fonts only
    SetBkMode .hDC_H, 1
    
    '----Create shape----
    BeginPath .hdc
        Rectangle .hdc, .mPoints(1).X, .mPoints(1).Y, .mPoints(3).X, .mPoints(3).Y
    EndPath .hdc
    
    mRGN = PathToRegion(.hdc)
    
    brush = CreateSolidBrush(mColor)
    SelectObject .hdc, brush
    FillRgn .hdc, mRGN, brush
    
    brush = CreateSolidBrush(hColor)
    SelectObject .hDC_H, brush
    FillRgn .hDC_H, mRGN, brush
      
    '----draw button lines----
    pen = CreatePen(0, 1, ms2color)
    RetVal = SelectObject(.hdc, pen)
    RetVal = LineTo(.hdc, .mPoints(1).X, .mPoints(1).Y)
    RetVal = LineTo(.hdc, .mPoints(2).X - 1, .mPoints(2).Y)
    RetVal = MoveToEx(.hdc, .mPoints(2).X - 2, .mPoints(2).Y + 1, prev)
    pen = CreatePen(0, 1, hs2Color)
    RetVal = SelectObject(.hdc, pen)
    RetVal = LineTo(.hdc, .mPoints(3).X - 2, .mPoints(3).Y - 2)
    RetVal = LineTo(.hdc, .mPoints(0).X, .mPoints(0).Y - 2)
    
    '----GRADIENT----
    MoveToEx .hdc, 0, 0, prev
    For I = 1 To .mH - 2
    If I Mod 2 <> 0 Then
    MoveToEx .hdc, 1, I, prev
    pen = CreatePen(0, 1, RGB(155 + (((255 - 155) / .mH) * I), 0, 0))
    RetVal = SelectObject(.hdc, pen)
    LineTo .hdc, .mW - 2, I
    Else
    MoveToEx .hdc, 1, I, prev
    pen = CreatePen(0, 1, RGB(0, 0, 0))
    RetVal = SelectObject(.hdc, pen)
    LineTo .hdc, .mW - 2, I
    End If
    Next

    MoveToEx .hDC_M, 0, 0, prev
    For I = 1 To .mH - 2
    If I Mod 2 <> 0 Then
    MoveToEx .hDC_M, 1, I, prev
    pen = CreatePen(0, 1, RGB(155 + (((255 - 155) / .mH) * I), 0, 0))
    RetVal = SelectObject(.hDC_M, pen)
    LineTo .hDC_M, .mW - 2, I
    Else
    MoveToEx .hDC_M, 1, I, prev
    pen = CreatePen(0, 1, RGB(255, 255, 255))
    RetVal = SelectObject(.hDC_M, pen)
    LineTo .hDC_M, .mW - 2, I
    End If
    Next


    '----draw button lines_hover----
    pen = CreatePen(0, 1, hs1Color)
    RetVal = SelectObject(.hDC_H, pen)
    RetVal = MoveToEx(.hDC_H, .mPoints(0).X, .mPoints(0).Y, prev)
    RetVal = LineTo(.hDC_H, .mPoints(1).X, .mPoints(1).Y)
    RetVal = LineTo(.hDC_H, .mPoints(2).X - 1, .mPoints(2).Y)
    RetVal = MoveToEx(.hDC_H, .mPoints(2).X - 2, .mPoints(2).Y + 1, prev)
    pen = CreatePen(0, 1, hs2Color)
    RetVal = SelectObject(.hDC_H, pen)
    RetVal = LineTo(.hDC_H, .mPoints(3).X - 2, .mPoints(3).Y - 2)
    RetVal = LineTo(.hDC_H, .mPoints(0).X, .mPoints(0).Y - 2)
    
    
End With

'----text----
With mRECT
    .Left = 5
    .Right = pObj.mPoints(1).X + pObj.mW
    .Top = pObj.mPoints(1).Y  '(mObj.mH / 2) - (ccA * 9)
    .Bottom = pObj.mPoints(1).X + pObj.mH
End With

font = CreateFontIndirect(curFont)

SelectObject pObj.hdc, font
SelectObject pObj.hDC_H, font

SetTextColor pObj.hdc, RGB(255, 255, 255)
tmpString = "____INVENTORY____"
DrawText pObj.hdc, tmpString, Len(tmpString), mRECT, 0


tmpString = "EXP : "
SetTextColor pObj.hdc, RGB(255, 255, 0)
mRECT.Top = mRECT.Top + (12)
DrawText pObj.hdc, tmpString, Len(tmpString), mRECT, 0

tmpString = "LEVEL : "
mRECT.Top = mRECT.Top + (12)
DrawText pObj.hdc, tmpString, Len(tmpString), mRECT, 0


BitBlt cHDC, pX, pY, pObj.mW, pObj.mH, pObj.hDC_M, 0, 0, vbSrcAnd
BitBlt cHDC, pX, pY, pObj.mW, pObj.mH, pObj.hdc, 0, 0, vbSrcPaint

'----delete unusable objects----
ReleaseDC 0, ScreenDC
DeleteObject picbmp
DeleteObject picbmpH
DeleteObject picbmpM
DeleteObject brush
DeleteObject pen

End Sub

Public Sub mCreateCircle(pX As Long, pY As Long, H_STRING As String)
anim = False
charT = BULLET
bObj.NOP = 3
ReDim bObj.mPoints(3)
absX = pX: absY = pY
UText = H_STRING

mColor = RGB(212, 208, 200) '13160660
ms1Color = RGB(255, 255, 255)
ms2color = RGB(180, 180, 180)
hColor = RGB(150, 0, 0) ' 38400
hs1Color = RGB(255, 200, 200)
hs2Color = RGB(150, 0, 0)

'----coordinates----
With bObj
    .mPoints(0).X = 0: .mPoints(0).Y = 3  'lower left corner
    .mPoints(1).X = 0: .mPoints(1).Y = 0 'upper left corner
    .mPoints(2).X = 3: .mPoints(2).Y = 0  'lower right corner
    .mPoints(3).X = 3: .mPoints(3).Y = 3  'upper right corner
    .mW = (.mPoints(2).X - .mPoints(0).X)
    .mH = (.mPoints(0).Y - .mPoints(2).Y)
    ScreenDC = GetDC(0)
    .hdc = CreateCompatibleDC(ScreenDC)
    .hDC_H = CreateCompatibleDC(ScreenDC)
    .hDC_M = CreateCompatibleDC(ScreenDC)
    picbmp = CreateCompatibleBitmap(ScreenDC, .mW, .mH)
    picbmpH = CreateCompatibleBitmap(ScreenDC, .mW, .mH)
    picbmpM = CreateCompatibleBitmap(ScreenDC, .mW, .mH)
    SelectObject .hdc, picbmp
    SelectObject .hDC_H, picbmpH
    SelectObject .hDC_M, picbmpM
'---------------------

    '----fill----
    SetBkColor .hdc, RGB(0, 0, 0)
    SetBkColor .hDC_H, RGB(0, 0, 0)
    SetBkColor .hDC_M, RGB(255, 255, 255)
    
    SetBkMode bObj.hdc, 1
    SetBkMode bObj.hDC_H, 1
    SetBkMode bObj.hDC_M, 0
    '------------
    
    '----create mask background----
    BeginPath .hDC_M
    Rectangle .hDC_M, .mPoints(1).X - 1, .mPoints(1).Y - 1, .mPoints(3).X + 1, .mPoints(3).Y + 1
    EndPath .hDC_M
    mRGN = PathToRegion(.hDC_M)
    brush = CreateSolidBrush(RGB(255, 255, 255))
    SelectObject .hDC_M, brush
    FillRgn .hDC_M, mRGN, brush
    '------------------------------
    
    '----Create shape----
    BeginPath .hdc
      Ellipse .hdc, .mPoints(1).X, .mPoints(1).Y, .mPoints(3).X, .mPoints(3).Y
    EndPath .hdc
    
    mRGN = PathToRegion(.hdc)
    
    brush = CreateSolidBrush(RGB(200, 0, 0))
    SelectObject .hdc, brush
    FillRgn .hdc, mRGN, brush
    
    brush = CreateSolidBrush(hColor)
    SelectObject .hDC_H, brush
    FillRgn .hDC_H, mRGN, brush
    '----------------------
    
    '----create mask shape----
    brush = CreateSolidBrush(vbBlack)
    SelectObject .hDC_M, brush
    FillRgn .hDC_M, mRGN, brush
    '--------------------------
    
    '----draw button lines----
            'down
    pen = CreatePen(0, 1, ms2color)
    RetVal = SelectObject(.hdc, pen)
    'Arc .hdc, .mPoints(1).x, .mPoints(1).y, .mPoints(3).x, .mPoints(3).y, .mPoints(0).x, .mPoints(0).y / 2, .mPoints(2).x, .mPoints(3).y / 2
    
            'up
    pen = CreatePen(0, 1, vbWhite)
    RetVal = SelectObject(.hdc, pen)
    SetArcDirection .hdc, 2
    'Arc .hdc, .mPoints(0).x, .mPoints(0).y, .mPoints(2).x, .mPoints(2).y, .mPoints(1).x, .mPoints(0).y / 2, .mPoints(2).x, .mPoints(3).y
    Ellipse .hdc, .mPoints(1).X, .mPoints(1).Y, .mPoints(3).X, .mPoints(3).Y
    '---------------------------
    
    '----draw button lines_hover----
    pen = CreatePen(0, 1, hs1Color)
    RetVal = SelectObject(.hDC_H, pen)
    Ellipse .hDC_H, .mPoints(1).X, .mPoints(1).Y, .mPoints(3).X, .mPoints(3).Y
    '-------------------------------
End With

'text
With mRECT
    .Left = 0
    .Right = bObj.mPoints(1).X + bObj.mW
    .Top = (bObj.mH / 2) - 9
    .Bottom = bObj.mPoints(1).X + bObj.mH
End With

font = CreateFontIndirect(curFont)

SelectObject bObj.hdc, font
SelectObject bObj.hDC_H, font

SetTextColor bObj.hdc, RGB(255, 255, 255)
DrawText bObj.hdc, H_STRING, Len(H_STRING), mRECT, 1

SetTextColor bObj.hDC_H, RGB(255, 255, 255)
DrawText bObj.hDC_H, H_STRING, Len(H_STRING), mRECT, 1

ReleaseDC 0, ScreenDC
DeleteObject picbmp
DeleteObject picbmpH
DeleteObject picbmpM
DeleteObject brush
DeleteObject pen

End Sub

Private Sub cmdLoadFile(fl As String)
Dim flOffset As Integer, TotAct As Integer, nFile As String
flOffset = 1
nFile = Mid(fl, 1, Len(fl) - 4) 'App.Path & "\walk"
Open nFile For Binary Access Read As #1
    Get #1, flOffset, cAction
Close #1

End Sub

Public Sub mMoveBullet(iobj As Integer)
With buObj(iobj)

If .animR Then

.oX = .lPath(.bCnt).X - (.mW / 2)
.oY = .lPath(.bCnt).Y - (.mH)

.bCnt = .bCnt + 15

If .bCnt >= UBound(.lPath()) Then .animR = False: .bCnt = UBound(.lPath())

End If

End With

End Sub

Public Sub mMoveEnemy(iobj As Integer)

With hObj(iobj)

If .animR Then

    .oX = .lPath(.bCnt).X - (.mW / 2)
    .oY = .lPath(.bCnt).Y - (.mH)
    .bCnt = .bCnt + 5
    If .bCnt >= UBound(.lPath()) Then .animR = False: .bCnt = UBound(.lPath())

End If

End With

End Sub

Public Sub mMoveT(iobj As Integer, sX As Long, sY As Long)

Static cX As Long, cY As Long

If cObj(iobj).animR Then

With cObj(iobj)
    
    cX = .oX
    cY = .oY

    cX = cX + sX
    cY = cY + sY
    .oX = cX
    .oY = cY
    
    If .oX < 0 Then .oX = 0
    If .oX > 500 Then .oX = 500

    If .oY < 30 Then .oY = 30
    If .oY > 400 - .mH(.F) Then .oY = 400 - .mH(.F)

    .F = .F + 1
    If .F > 1 Then .F = 0

End With

End If
End Sub


Public Sub CollisionDetection()

Dim cd1 As Integer, cd2 As Integer, bi As Integer

If craftLife = 0 Or gFinale Then Exit Sub
'For cd1 = LBound(cObj) To UBound(cObj)
    
    For cd2 = LBound(hObj) To UBound(hObj) 'Step -1

        'If GamePause = False And
        If HeroKill = False Then '2
    
    With hObj(cd2)
        .oRECT.Left = .oX
        .oRECT.Right = .oX + .mW
        .oRECT.Bottom = .oY + .mH
        .oRECT.Top = .oY
    End With
    
    With cObj(HERO)
        .oRECT(.F).Left = .oX
        .oRECT(.F).Right = .oX + .mW(.F)
        .oRECT(.F).Bottom = .oY + .mH(.F)
        .oRECT(.F).Top = .oY
    End With
                
                If IntersectRect(tempRect, cObj(HERO).oRECT(cObj(HERO).F), hObj(cd2).oRECT) Then 'Or _

                intL = True
                
                cObj(HERO).animR = False
                cObj(HERO).F = 0
                cObj(HERO) = emptyObjM
                Erase buObj
                DuplicateObjS cd2
                setCoordsS hObj(cd2).oX, hObj(cd2).oY, cd2
                sObj(cd2).animR = True
                fireMode = False
                HeroKill = True
                'sBOOM.PlayWave
                sndPlaySound App.Path & "\sounds\boom.wav", SND_ASYNC Or SND_NODEFAULT
                dT1 = GetTickCount
                
                Exit Sub
            
            End If
        End If '2
        
        mMoveEnemy cd2
    
    Next cd2
    
    intL = False

'Next cd1


If fireMode Then

For cd1 = LBound(buObj) To UBound(buObj)
'For cd1 = 0 To 99
mMoveBullet cd1
    For cd2 = LBound(hObj) To UBound(hObj)

        'If GamePause = False Or HeroKill = False Then '3
        If HeroKill = False Then '3
        If buObj(cd1).animR Then  '2
    
    With buObj(cd1)
        .oRECT.Left = .oX
        .oRECT.Right = .oX + .mW
        .oRECT.Bottom = .oY + .mH
        .oRECT.Top = .oY
    End With

    With hObj(cd2)
        .oRECT.Left = .oX
        .oRECT.Right = .oX + .mW
        .oRECT.Bottom = .oY + .mH
        .oRECT.Top = .oY
    End With
    
                If IntersectRect(tempRect, buObj(cd1).oRECT, hObj(cd2).oRECT) Then
                    sBOOM.StopWave
                    DuplicateObjS cd2
                    setCoordsS hObj(cd2).oX, hObj(cd2).oY, cd2
                    sObj(cd2).animR = True
                    buObj(cd1).animR = False
                    buObj(cd1) = emptyObj
                    hObj(cd2) = emptyObj
                    Hits = Hits + 1
                    'sBOOM.PlayWave
                    sndPlaySound App.Path & "\sounds\boom.wav", SND_ASYNC Or SND_NODEFAULT
                Exit Sub
            End If
        End If '2
        
        End If '3
    
    Next cd2



Next cd1

End If
End Sub



Public Sub mAnim(iobj As Integer, maxFr As Integer)

If sObj(iobj).animR Then

With sObj(iobj)

.F = .F + 1
If .F >= maxFr Then .animR = False: .F = 0


End With

End If

End Sub

Public Sub sDraw(cHDC As Long, str As Long)
With tObj
    'BitBlt cHDC, 320 - (.mW / 2), 100, .mW, .mH, .hDC_M, 0, 0, vbSrcAnd
    'BitBlt cHDC, 320 - (.mW / 2), 100, .mW, .mH, .hdc, 0, 0, vbSrcPaint
    StretchBlt cHDC, (320 - (.mW / 2)) + str / 2, 100, .mW - str, .mH - str, .hDC_M, 0, 0, .mW, .mH, vbSrcAnd
    StretchBlt cHDC, (320 - (.mW / 2)) + str / 2, 100, .mW - str, .mH - str, .hdc, 0, 0, .mW, .mH, vbSrcPaint
End With

End Sub


Public Sub reDraw(cHDC As Long)
Dim r As Integer
If craftLife = 0 Then Exit Sub
If gFinale Then GoTo FinaleSub
totalSC = (Hits * 50)

mDrawText cHDC, 5, 5, "SCORE : " & totalSC & "    LIFE : " & craftLife & "   HITS :" & Hits & "/" & AllEnemies & "/" & HitsPerLevel & "   LEVEL : " & curLevel & "    FPS: " & IntFrames & " / " & cDelay & " ms", 14, RGB(255, 255, 0)

For r = LBound(hObj) To UBound(hObj)

With hObj(r)
    BitBlt cHDC, .oX, .oY, .mW, .mH, .hDC_M, 0, 0, vbSrcAnd
    BitBlt cHDC, .oX, .oY, .mW, .mH, .hdc, 0, 0, vbSrcPaint
    
    If DebugM Then Rectangle cHDC, .oRECT.Left, .oRECT.Top, .oRECT.Right, .oRECT.Bottom

End With
Next r

'For r = LBound(cObj) To UBound(cObj)

If HeroKill = False Then

With cObj(HERO)
      
    BitBlt cHDC, .oX + .bX(.F), .oY + .bY(.F), .mW(.F), .mH(.F), .hDC_M(.F), 0, 0, vbSrcAnd
    BitBlt cHDC, .oX + .bX(.F), .oY + .bY(.F), .mW(.F), .mH(.F), .hdc(.F), 0, 0, vbSrcPaint

    If DebugM Then Rectangle cHDC, .oRECT(HERO).Left, .oRECT(HERO).Top, _
        .oRECT(HERO).Right, .oRECT(HERO).Bottom
 
End With

End If
'Next r

If fireMode Then
For r = LBound(buObj) To UBound(buObj)

With buObj(r)
      
    BitBlt cHDC, .oX, .oY, .mW, .mH, .hDC_M, 0, 0, vbSrcAnd
    BitBlt cHDC, .oX, .oY, .mW, .mH, .hdc, 0, 0, vbSrcPaint
    

    If DebugM Then Rectangle cHDC, .oRECT.Left, .oRECT.Top, .oRECT.Right, .oRECT.Bottom

End With

Next r

End If

For r = 0 To UBound(sObj)

With sObj(r)
    
    If .animR Then
        mAnim r, .NOPMu
        BitBlt cHDC, .oX, .oY, .mW(.F), .mH(.F), .hDC_M(.F), 0, 0, vbSrcAnd
        BitBlt cHDC, .oX, .oY, .mW(.F), .mH(.F), .hdc(.F), 0, 0, vbSrcPaint
    End If

End With

Next r


If GamePause Then mDrawText cHDC, 250, 150, "LEVEL " & curLevel & " COMPLETED", 20, RGB(255, 255, 0)
If HeroKill Then mDrawText cHDC, 250, 150, "YOU LOST...", 20, RGB(0, 255, 0)
Exit Sub
FinaleSub:
fT2 = GetTickCount
If gFinale Then mDrawText cHDC, 25, 150, "YOU HAVE FINISHED STAGE 1 AND THIS DEMO SUCCESFULL", 20, RGB(0, 255, 0)
If fT2 - fT1 >= 10000 Then GameOver
End Sub

Private Sub LineB(cHDC As Long, mX As Long, mY As Long, iobj As Integer)
With buObj(iobj)
.bCnt = 1
drawLine cHDC, cObj(HERO).oX + cObj(HERO).mW(cObj(HERO).F), _
    (cObj(HERO).oY + (cObj(HERO).mH(cObj(HERO).F) / 2)), (mX), (mY), iobj, False
.animR = True
End With
End Sub


Public Sub LineH(cHDC As Long, mX As Long, mY As Long, iobj As Integer, Optional mReset As Boolean, Optional nX As Long, Optional nY As Long)
If mReset Then

With hObj(iobj)
.bCnt = 1
drawLine cHDC, .oX, (.oY), (mX), (mY), iobj, True
drawLine cHDC, (mX), (mY), nX, nY, iobj, True, mReset
.animR = True
End With

Else

With hObj(iobj)
.bCnt = 1
drawLine cHDC, .oX, (.oY), (mX), (mY), iobj, True, mReset
.animR = True
End With

End If
End Sub

Public Sub LineM(cHDC As Long, mX As Long, mY As Long, iobj As Integer)
Dim sAngle As Double

With cObj(iobj)
sAngle = InvTanForTheta((mX - .oX) / (mY - .oY))
gCnt = 1

If mY > .oY Then

Select Case sAngle

Case -15 To 15
    .F = 3: cn = 3 'Down

Case -80 To -15
    .F = 12: cn = 12  'downleft

Case 15 To 80
    .F = 15: cn = 15  'downright

Case -90 To -80
    .F = 9: cn = 9  'left

Case 80 To 90
    .F = 6: cn = 6  'right

End Select

ElseIf mY < .oY Then

.F = 0: cn = 0

Select Case sAngle

Case -15 To 15
    .F = 0: cn = 0  'up

Case -80 To -30
    .F = 18: cn = 18  'upright

Case 15 To 80
    .F = 21: cn = 21  'upleft

Case 80 To 90
    .F = 9: cn = 9  'left

Case -90 To -80
    .F = 6: cn = 6  'right

End Select

End If

If .Class = HERO Then
drawLine cHDC, .oX + (gObj.mW(.F) / 2), .oY + (gObj.mH(.F)), (mX), (mY), iobj, True
End If

End With
End Sub

Public Sub mDrawText(cHDC As Long, tX As Long, tY As Long, cTEXT As String, fSize As Integer, mColor As Long, Optional wFormat As Long)
createFONT cHDC, fSize
'----TEXT----
SetTextColor cHDC, mColor
SetBkMode cHDC, 1
With mRECT
    .Left = tX
    .Right = tX + (9 * fSize * Len(cTEXT))
    .Top = tY
    .Bottom = tY + (fSize * 12)
    'RetVal = Rectangle(cHDC, .Left, .Top, .Right, .Bottom)
End With

    RetVal = DrawText(cHDC, cTEXT, Len(cTEXT), mRECT, wFormat)
'----END TEXT----
End Sub


Friend Sub click_event(cHDC As Long)

If HeroKill Then
    dT2 = GetTickCount
    If dT2 - dT1 > 4000 Then gReset
End If
       
If craftLife = 0 Then Exit Sub


eT2 = GetTickCount

If Hits < HitsPerLevel Then
    If eT2 - eT1 > TIMEFORENEMIES Then
        newEnemy cHDC
        eT1 = GetTickCount
    End If

Else
    If GamePause = False Then
        GamePause = True
        pT1 = GetTickCount
    End If
End If

If GamePause Then
    pT2 = GetTickCount
    If pT2 - pT1 > 1000 Then
    gResetL
    End If
End If



If HeroKill = False Then 'HeroKill

With cObj(HERO)

If SupJoystic = 0 Then 'Joystick

GetKeyboardInput myKey


If myKey.btnUp Then
    
    mMoveT 0, 0, -mDist
    
End If

If myKey.btnDown Then
    
    mMoveT 0, 0, mDist

End If
If myKey.btnLeft Then
    
    mMoveT 0, -mDist, 0

End If
If myKey.btnRight Then
    
    mMoveT 0, mDist, 0

End If

If myKey.btnControl Then
    bT2 = GetTickCount
        If bT2 - bT1 > 40 Then
        
            Shoot cHDC
        End If
    bT1 = GetTickCount
End If

cObj(HERO).animR = True

'JOYSTICK
Else

joyGetPos JOYSTICKID1, mJoy

If mJoy.wYpos = 0 Then
    
    mMoveT 0, 0, -mDist
    cObj(HERO).animR = True

End If
If mJoy.wYpos > 32800 Then
    
    mMoveT 0, 0, mDist
    cObj(HERO).animR = True

End If
If mJoy.wXpos = 0 Then
    
    mMoveT 0, -mDist, 0
    cObj(HERO).animR = True

End If
If mJoy.wXpos > 32800 Then
    
    mMoveT 0, mDist, 0
    cObj(HERO).animR = True

End If

If mJoy.wButtons = 1 Then
    bT2 = GetTickCount
        If bT2 - bT1 > 40 Then
        
            Shoot cHDC
        End If
    bT1 = GetTickCount
End If
 

End If 'JOYSTICK




End With

End If 'Herokill

End Sub



Private Function InvTanForTheta(ByVal Number As Double) As Double
On Error GoTo PROC_Err

Dim theta As Double

theta = Atn(Number) * RadToDegFuncX

InvTanForTheta = theta


PROC_Exit:
Exit Function
PROC_Err:
InvTanForTheta = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Function

Public Function LoadGraphicDC(sFileName As String) As Long
On Error Resume Next
Dim LoadGraphicDCTEMP As Long

LoadGraphicDCTEMP = CreateCompatibleDC(GetDC(0))

SelectObject LoadGraphicDCTEMP, LoadPicture(sFileName)

LoadGraphicDC = LoadGraphicDCTEMP

End Function
Public Sub LoadScreen(hdc As Long, hH As Long)
On Error Resume Next
Dim LoadGraphicDCTEMP As Long
Dim mBITMAP As Long

LoadGraphicDCTEMP = CreateCompatibleDC(GetDC(0))
mBITMAP = CreateCompatibleBitmap(GetDC(0), 640, hH)
SelectObject LoadGraphicDCTEMP, mBITMAP
hdc = LoadGraphicDCTEMP
DeleteObject mBITMAP
End Sub

Public Sub createFONT(cHDC As Long, ccA As Integer)
'ccA = 1
With curFont
    .lfHeight = -(ccA)
    .lfWidth = 0
    .lfEscapement = 0
    .lfOrientation = 0
    .lfWeight = 400
    .lfItalic = 0
    .lfUnderline = 0
    .lfStrikeOut = 0
    .lfCharSet = DEFAULT_CHARSET
    .lfOutPrecision = 2
    .lfClipPrecision = 1
    .lfQuality = 0
    .lfPitchAndFamily = 32
    .lfFaceName = 1
End With
    
    font = CreateFontIndirect(curFont)
    
    SelectObject cHDC, font
    SetTextColor cHDC, RGB(0, 0, 0)
    DeleteObject font
End Sub

Public Sub setCoordsH(mX As Long, mY As Long, curSprite As Integer, sprtName As String)
hObj(curSprite).oX = mX: hObj(curSprite).oY = mY
hObj(curSprite).oName = sprtName
End Sub

Public Sub setCoordsB(mX As Long, mY As Long, curSprite As Integer)
buObj(curSprite).oX = mX: buObj(curSprite).oY = mY

End Sub

Public Sub setCoordsS(mX As Long, mY As Long, curSprite As Integer)
sObj(curSprite).oX = mX: sObj(curSprite).oY = mY
End Sub

Public Sub setCoords(mX As Long, mY As Long, curSprite As Integer, sprtName As String, sprtClass As Integer)
cObj(curSprite).oX = mX: cObj(curSprite).oY = mY
cObj(curSprite).oName = sprtName
cObj(curSprite).Class = sprtClass
End Sub

Private Sub Class_Terminate()
If LoadedSprites Then

ReleaseDC 0, ScreenDC
DeleteObject picbmp
DeleteObject picbmpH
DeleteObject picbmpM

For I = 0 To UBound(gObj.hdc)
DeleteDC gObj.hdc(I)
DeleteDC gObj.hDC_H(I)
DeleteDC gObj.hDC_M(I)
Next

For I = 0 To UBound(gCraftObj.hdc)
DeleteDC gCraftObj.hdc(I)
DeleteDC gCraftObj.hDC_H(I)
DeleteDC gCraftObj.hDC_M(I)
Next I

For I = 0 To UBound(gSmokeObj.hdc)
DeleteDC gSmokeObj.hdc(I)
DeleteDC gSmokeObj.hDC_H(I)
DeleteDC gSmokeObj.hDC_M(I)
Next I

Erase sObj

DeleteDC mObj.hdc
DeleteDC mObj.hDC_H
DeleteDC mObj.hDC_M

DeleteDC pObj.hdc
DeleteDC pObj.hDC_H
DeleteDC pObj.hDC_M

DeleteDC bObj.hdc
DeleteDC bObj.hDC_H
DeleteDC bObj.hDC_M

DeleteObject brush
DeleteObject pen

End If
End Sub

Public Sub Shoot(cHDC As Long)

If HeroKill Then cObj(HERO).oY = 223: tBull = 0: Exit Sub
    fireMode = True
    tBull = tBull + 1
    Allbul = Allbul + 1
    If tBull >= UBound(buObj) Then tBull = 0
    DuplicateObjB tBull
    LineB cHDC, 700, cObj(HERO).oY, tBull
    sSHOOT.PlayWave

End Sub

Public Sub newEnemy(cHDC As Long)

Dim rY As Long

If HeroKill Then curEn = 0: Exit Sub

DuplicateObjH curEn


'Select Case curEn Mod 5
'Case 0


'rY = 200

'setCoordsH 640, rY, curEn, "ENEMY"

'Randomize
'LineH cHDC, Int((400 * Rnd) + 200), 50, curEn, True, -50, cObj(HERO).oY + 8


'Case Is = 4


'rY = 200

'setCoordsH 640, rY, curEn, "ENEMY"

'Randomize
'LineH cHDC, Int((400 * Rnd) + 200), 400, curEn, True, -50, cObj(HERO).oY + 8

'Case 1 To 3


Randomize

rY = Int((400 * Rnd) + 30)

setCoordsH 640, rY, curEn, "ENEMY"

LineH cHDC, cObj(HERO).oX, cObj(HERO).oY + cObj(HERO).mH(cObj(HERO).F), curEn, True, -50, cObj(HERO).oY


'End Select

curEn = curEn + 1

AllEnemies = AllEnemies + 1


End Sub

Private Sub gReset()
DuplicateObj 0
setCoords 100, 258, HERO, "HERO", HERO
HeroKill = False
craftLife = craftLife - 1
If craftLife = 0 Then GameOver
End Sub

Private Sub GameOver()

fCancel = True: fEnter = False

Dim EOFl As Long

If Dir(App.Path & "\hiscores.dat") <> "" Then
    EOFl = FileLen(App.Path & "\hiscores.dat")
End If
    
    Open App.Path & "\hiscores.dat" For Binary Access Write As #1

        With mScores
            .hScore = totalSC
            .hLevel = curLevel
            .hName = PlayerName
            .hHits = Hits
            .hTotalHits = AllEnemies
            .hDATE = Date
        End With
        Put #1, EOFl + 1, mScores
    Close #1

Erase hObj
gameRun = False

End Sub
Private Sub gResetL()

Hits = HitsPerLevel

curLevel = curLevel + 1

If curLevel = 8 Then fT1 = GetTickCount: EndCondition

HitsPerLevel = (Hits) + curLevel * 100

GameCondition curLevel, dLEVEL

GamePause = False
End Sub

Private Sub EndCondition()
Erase hObj
gFinale = True
End Sub

Private Function GetKeyboardInput(Keyboard As KeyboardInput)

With Keyboard
    .btnDown = GetAsyncKeyState(vbKeyDown)
    .btnUp = GetAsyncKeyState(vbKeyUp)
    .btnRight = GetAsyncKeyState(vbKeyRight)
    .btnLeft = GetAsyncKeyState(vbKeyLeft)
    .btnA = GetAsyncKeyState(vbKeyA)
    .btnB = GetAsyncKeyState(vbKeyB)
    .btnC = GetAsyncKeyState(vbKeyC)
    .btnD = GetAsyncKeyState(vbKeyD)
    .btnE = GetAsyncKeyState(vbKeyE)
    .btnF = GetAsyncKeyState(vbKeyF)
    .btnG = GetAsyncKeyState(vbKeyG)
    .btnH = GetAsyncKeyState(vbKeyH)
    .btnI = GetAsyncKeyState(vbKeyI)
    .btnJ = GetAsyncKeyState(vbKeyJ)
    .btnK = GetAsyncKeyState(vbKeyK)
    .btnL = GetAsyncKeyState(vbKeyL)
    .btnM = GetAsyncKeyState(vbKeyM)
    .btnN = GetAsyncKeyState(vbKeyN)
    .btnO = GetAsyncKeyState(vbKeyO)
    .btnP = GetAsyncKeyState(vbKeyP)
    .btnQ = GetAsyncKeyState(vbKeyQ)
    .btnR = GetAsyncKeyState(vbKeyR)
    .btnS = GetAsyncKeyState(vbKeyS)
    .btnT = GetAsyncKeyState(vbKeyT)
    .btnU = GetAsyncKeyState(vbKeyU)
    .btnV = GetAsyncKeyState(vbKeyV)
    .btnW = GetAsyncKeyState(vbKeyW)
    .btnX = GetAsyncKeyState(vbKeyX)
    .btnY = GetAsyncKeyState(vbKeyY)
    .btnZ = GetAsyncKeyState(vbKeyZ)
    .btnControl = GetAsyncKeyState(vbKeyControl)
    .btnReturn = GetAsyncKeyState(vbKeyReturn)
End With

End Function



Friend Sub Starfield(cHDC As Long)
Dim I As Long
Static sLoop As Long

If sLoop = 0 Then
Randomize
'Generate the 100 stars
For I = LBound(Stars) To UBound(Stars)
    
    Stars(I).X = 1280 * Rnd + 640
    Stars(I).Y = 400 * Rnd + 1
    Stars(I).Speed = MaxSpeed * Rnd + 1
Next I
sLoop = 1
End If

For I = 0 To UBound(Stars)
    
    'Move the star
    Stars(I).X = (Stars(I).X Mod 1280) - Stars(I).Speed
    'Relocate the X position
    If Stars(I).X < 0 Then
      'Stars(I).Y = 400 * Rnd + 1
      Stars(I).X = 1280 * Rnd + 640
      Stars(I).Speed = MaxSpeed * Rnd + 1
    End If
    'Draw the star
    SetPixelV cHDC, Stars(I).X, Stars(I).Y, vbWhite
Next I

End Sub

Public Sub SolidRect(cHDC As Long, sX1 As Long, sY1 As Long, sX2 As Long, sY2 As Long)
Dim sBrush As Long, sRGN As Long
    
    BeginPath cHDC
    Rectangle cHDC, sX1, sY1, sX2, sY2
    EndPath cHDC
    sRGN = PathToRegion(cHDC)
    sBrush = CreateSolidBrush(RGB(0, 0, 0))
    SelectObject cHDC, sBrush
    FillRgn cHDC, sRGN, sBrush
End Sub

