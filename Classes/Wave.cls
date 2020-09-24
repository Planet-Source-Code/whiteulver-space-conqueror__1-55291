VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Wave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'-Start Direct x------------------------------------------
Private Type WAVETYPE '36 byte
    strHead As String * 4
    lngFile As Long
    strHead_RIFF As String * 4
    strFormatID As String * 4
    lngChunkSize As Long
    intFormat As Integer
    intChannels As Integer
    lngSamplesPerSec As Long
    lngAvgBytesPerSec As Long
    intBlockAlign As Integer
    intBitsPerSample As Integer
End Type
'DirectX Variables
Private DX As DirectX7
Private DxSound As DirectSound
Private jBuffer As DirectSoundBuffer

'-Start zlib-----------------------------------------------
Private Declare Function Compress Lib "zlib" Alias "compress2" (Dest As Any, destLen As Any, Src As Any, ByVal SrcLen As Long, ByVal Level As Long) As Long
Private Declare Function Decompress Lib "zlib" Alias "uncompress" (Dest As Any, destLen As Any, Src As Any, ByVal SrcLen As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Const Z_OK              As Long = 0
Private Const Z_STREAM_END      As Long = 1
Private Const Z_NEED_DICT       As Long = 2
Private Const Z_ERRNO           As Long = -1
Private Const Z_STREAM_ERROR    As Long = -2
Private Const Z_DATA_ERROR      As Long = -3
Private Const Z_MEM_ERROR       As Long = -4
Private Const Z_BUF_ERROR       As Long = -5
Private Const Z_VERSION_ERROR   As Long = -6

Private Declare Function zlibVersion Lib "zlib.dll" () As String
'-Other Variables-----------------------------------------------

Private Type chunk
    Chuky As String * 4
    DataLen As Long
End Type
Dim SC As String * 128
Dim SCCH As Long
Dim OLEN As Long
Dim CLEN As Long
Dim CMP As Boolean
Dim wName As String
Dim lName As String

Dim FM As WAVETYPE
Dim DM As chunk
Public Sub InitSound(sFile As String, mForm As Form)
Initialize mForm

'On Error Resume Next
    Dim udtBufferDesc As DSBUFFERDESC
    Dim udtFormat As WAVEFORMATEX
    Dim Buf() As Byte
    
    On Error Resume Next
    lName = sFile
    
    Open lName For Binary As #1
    
    Get #1, , FM
    Get #1, , DM
    
    'If CMP Then
    '    Get #1, , OLEN
    '    Get #1, , CLEN
    '    ReDim Buf(CLEN)
    '    Get #1, , Buf
    '    DecompressArray Buf, OLEN
        
    'Else
        ReDim Buf(DM.DataLen - 1)
        Get #1, , Buf
        
    'End If
    
    
    
    With udtFormat
        .nFormatTag = FM.intFormat
        .nChannels = FM.intChannels
        .lSamplesPerSec = FM.lngSamplesPerSec
        .nBitsPerSample = FM.intBitsPerSample
        .nBlockAlign = FM.intBlockAlign
        .lAvgBytesPerSec = FM.lngAvgBytesPerSec
    End With
    'ChunkLen(i) = DM.DataLen
    
    
    udtBufferDesc.lBufferBytes = DM.DataLen
    Set jBuffer = DxSound.CreateSoundBuffer(udtBufferDesc, udtFormat)
    jBuffer.WriteBuffer 0, DM.DataLen, Buf(0), DSBLOCK_ENTIREBUFFER
    
    
    
    Erase Buf
    Close #1
End Sub
Public Sub PlayWave()

    jBuffer.Play DSBPLAY_DEFAULT
End Sub

Public Sub StopWave()
    jBuffer.Stop
End Sub
'-Direct X functions-----------------------------------------------
Private Sub Initialize(frmInit As Form)

    'Initialize DirectSound
    Set DX = New DirectX7
    Set DxSound = DX.DirectSoundCreate("")
    
    'Set the DirectSound object's cooperative level (Priority gives us sole control)
    DxSound.SetCooperativeLevel frmInit.hwnd, DSSCL_NORMAL ' DSSCL_PRIORITY
        
End Sub

Public Sub Terminate()

    'Terminate all
   
    
    Set jBuffer = Nothing
    Set DxSound = Nothing
    Set DX = Nothing

End Sub
'-zlib Functions----------------------------------------------------
Function CompressArray(TheData() As Byte, CompressionLevel As Integer) As Long
  Dim result As Long
  Dim BufferSize As Long
  Dim TempBuffer() As Byte

  BufferSize = UBound(TheData) + 1
  BufferSize = BufferSize + (BufferSize * 0.01) + 12
  ReDim TempBuffer(BufferSize)
  result = Compress(TempBuffer(0), BufferSize, TheData(0), UBound(TheData) + 1, CompressionLevel)
  ReDim Preserve TheData(BufferSize - 1)
  CopyMemory TheData(0), TempBuffer(0), BufferSize
  Erase TempBuffer
  CompressArray = result
End Function

Function DecompressArray(TheData() As Byte, OrigSize As Long) As Long
  Dim result As Long
  Dim BufferSize As Long
  Dim TempBuffer() As Byte

  BufferSize = OrigSize
  BufferSize = BufferSize + (BufferSize * 0.01) + 12
  ReDim TempBuffer(BufferSize)
  result = Decompress(TempBuffer(0), BufferSize, TheData(0), UBound(TheData) + 1)
  ReDim Preserve TheData(BufferSize - 1)
  CopyMemory TheData(0), TempBuffer(0), BufferSize
  DecompressArray = result
End Function

