Attribute VB_Name = "JPGSpeichern"
Option Explicit

Private Enum IJLERR
  IJL_OK = 0
  IJL_INTERRUPT_OK = 1
  IJL_ROI_OK = 2
  IJL_EXCEPTION_DETECTED = -1
  IJL_INVALID_ENCODER = -2
  IJL_UNSUPPORTED_SUBSAMPLING = -3
  IJL_UNSUPPORTED_BYTES_PER_PIXEL = -4
  IJL_MEMORY_ERROR = -5
  IJL_BAD_HUFFMAN_TABLE = -6
  IJL_BAD_QUANT_TABLE = -7
  IJL_INVALID_JPEG_PROPERTIES = -8
  IJL_ERR_FILECLOSE = -9
  IJL_INVALID_FILENAME = -10
  IJL_ERROR_EOF = -11
  IJL_PROG_NOT_SUPPORTED = -12
  IJL_ERR_NOT_JPEG = -13
  IJL_ERR_COMP = -14
  IJL_ERR_SOF = -15
  IJL_ERR_DNL = -16
  IJL_ERR_NO_HUF = -17
  IJL_ERR_NO_QUAN = -18
  IJL_ERR_NO_FRAME = -19
  IJL_ERR_MULT_FRAME = -20
  IJL_ERR_DATA = -21
  IJL_ERR_NO_IMAGE = -22
  IJL_FILE_ERROR = -23
  IJL_INTERNAL_ERROR = -24
  IJL_BAD_RST_MARKER = -25
  IJL_THUMBNAIL_DIB_TOO_SMALL = -26
  IJL_THUMBNAIL_DIB_WRONG_COLOR = -27
  IJL_RESERVED = -99
End Enum

Private Enum IJLIOTYPE
  IJL_SETUP = -1&
  IJL_JFILE_READPARAMS = 0&
  IJL_JBUFF_READPARAMS = 1&
  IJL_JFILE_READWHOLEIMAGE = 2&
  IJL_JBUFF_READWHOLEIMAGE = 3&
  IJL_JFILE_READHEADER = 4&
  IJL_JBUFF_READHEADER = 5&
  IJL_JFILE_READENTROPY = 6&
  IJL_JBUFF_READENTROPY = 7&
  IJL_JFILE_WRITEWHOLEIMAGE = 8&
  IJL_JBUFF_WRITEWHOLEIMAGE = 9&
  IJL_JFILE_WRITEHEADER = 10&
  IJL_JBUFF_WRITEHEADER = 11&
  IJL_JFILE_WRITEENTROPY = 12&
  IJL_JBUFF_WRITEENTROPY = 13&
  IJL_JFILE_READONEHALF = 14&
  IJL_JBUFF_READONEHALF = 15&
  IJL_JFILE_READONEQUARTER = 16&
  IJL_JBUFF_READONEQUARTER = 17&
  IJL_JFILE_READONEEIGHTH = 18&
  IJL_JBUFF_READONEEIGHTH = 19&
  IJL_JFILE_READTHUMBNAIL = 20&
  IJL_JBUFF_READTHUMBNAIL = 21&

End Enum

Private Type JPEG_CORE_PROPERTIES_VB
  UseJPEGPROPERTIES As Long
  DIBBytes As Long
  DIBWidth As Long
  DIBHeight As Long
  DIBPadBytes As Long
  DIBChannels As Long
  DIBColor As Long
  DIBSubsampling As Long
  jpgFile As Long
  jpgBytes As Long
  jpgSizeBytes As Long
  jpgWidth As Long
  jpgHeight As Long
  jpgChannels As Long
  jpgColor As Long
  jpgSubsampling As Long
  jpgThumbWidth As Long
  jpgThumbHeight As Long
  cconversion_reqd As Long
  upsampling_reqd As Long
  jquality As Long
  jprops(0 To 19999) As Byte
End Type


Private Declare Function ijlInit Lib "U:\Programmierung\ijl16.dll" _
        (jcprops As Any) As Long
        
Private Declare Function ijlFree Lib "ijl16.dll" _
        (jcprops As Any) As Long
        
Private Declare Function ijlRead Lib "ijl16.dll" _
        (jcprops As Any, ByVal ioType As Long) As Long
        
Private Declare Function ijlWrite Lib "ijl16.dll" _
        (jcprops As Any, ByVal ioType As Long) As Long
        
Private Declare Function ijlGetLibVersion Lib _
        "ijl16.dll" () As Long
        
Private Declare Function ijlGetErrorString Lib _
        "ijl16.dll" (ByVal code As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias _
        "RtlMoveMemory" (lpvDest As Any, lpvSource As _
        Any, ByVal cbCopy As Long)
        
Private Declare Function GlobalAlloc Lib "kernel32" _
        (ByVal wFlags As Long, ByVal dwBytes As Long) _
        As Long
        
Private Declare Function GlobalFree Lib "kernel32" _
        (ByVal hMem As Long) As Long
        
Private Declare Function GlobalLock Lib "kernel32" _
       (ByVal hMem As Long) As Long
       
Private Declare Function GlobalUnlock Lib "kernel32" _
        (ByVal hMem As Long) As Long
        
Public Const GMEM_DDESHARE = &H2000
Public Const GMEM_DISCARDABLE = &H100
Public Const GMEM_DISCARDED = &H4000
Public Const GMEM_FIXED = &H0
Public Const GMEM_INVALID_HANDLE = &H8000
Public Const GMEM_LOCKCOUNT = &HFF
Public Const GMEM_MODIFY = &H80
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_NOCOMPACT = &H10
Public Const GMEM_NODISCARD = &H20
Public Const GMEM_NOT_BANKED = &H1000
Public Const GMEM_NOTIFY = &H4000
Public Const GMEM_SHARE = &H2000
Public Const GMEM_VALID_FLAGS = &H7F72
Public Const GMEM_ZEROINIT = &H40
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
        
Public Function Loadjpg(ByRef cDib As cDIBSection, _
                        ByVal sFile$) As Boolean

Dim tJ As JPEG_CORE_PROPERTIES_VB
Dim bFile() As Byte
Dim lR As Long
Dim lPtr As Long
Dim ljpgWidth As Long, ljpgHeight As Long
  
  lR = ijlInit(tJ)
  If lR = IJL_OK Then
    bFile = StrConv(sFile, vbFromUnicode)
    ReDim Preserve bFile(0 To UBound(bFile) + 1) As Byte
    bFile(UBound(bFile)) = 0
    lPtr = VarPtr(bFile(0))
    CopyMemory tJ.jpgFile, lPtr, 4
  
    lR = ijlRead(tJ, IJL_JFILE_READPARAMS)
    If lR <> IJL_OK Then
      MsgBox "Failed to read jpg"
    Else
      ljpgWidth = tJ.jpgWidth
      ljpgHeight = tJ.jpgHeight
  
      If cDib.Create(ljpgWidth, ljpgHeight) Then
        tJ.DIBWidth = ljpgWidth
        tJ.DIBPadBytes = cDib.BytesPerScanLine - ljpgWidth * 3
        tJ.DIBHeight = -ljpgHeight
        tJ.DIBChannels = 3&
        tJ.DIBBytes = cDib.DIBSectionBitsPtr
  
        lR = ijlRead(tJ, IJL_JFILE_READWHOLEIMAGE)
        If lR = IJL_OK Then
           Loadjpg = True
        Else
           MsgBox "Cannot read Image Data FROM file.", vbExclamation
        End If
       End If
    End If
    ijlFree tJ
  Else
    MsgBox "Failed to initialise the IJL library: " & lR, vbExclamation
  End If
End Function

Public Function Savejpg(ByRef cDib As cDIBSection, ByVal sFile$, _
                        Optional ByVal lQuality As Long = 90) As Boolean

  Dim tJ As JPEG_CORE_PROPERTIES_VB
  Dim bFile() As Byte
  Dim lPtr As Long
  Dim lR As Long
  On Error GoTo fehler
    lR = ijlInit(tJ)
    If lR = IJL_OK Then
      tJ.DIBWidth = cDib.Width
      tJ.DIBHeight = -cDib.Height
      tJ.DIBBytes = cDib.DIBSectionBitsPtr
      tJ.DIBPadBytes = cDib.BytesPerScanLine - cDib.Width * 3

      bFile = StrConv(sFile, vbFromUnicode)
      ReDim Preserve bFile(0 To UBound(bFile) + 1) As Byte
      bFile(UBound(bFile)) = 0
      lPtr = VarPtr(bFile(0))
      CopyMemory tJ.jpgFile, lPtr, 4
      tJ.jpgWidth = cDib.Width
      tJ.jpgHeight = cDib.Height
      tJ.jquality = lQuality

      lR = ijlWrite(tJ, IJL_JFILE_WRITEWHOLEIMAGE)
      If lR = IJL_OK Then
         Savejpg = True
      Else
         MsgBox "Failed to save to jpg", vbExclamation
      End If
      ijlFree tJ
    Else
      MsgBox "Failed to initialise the IJL library: " & lR, vbExclamation
    End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Savejpg/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function


