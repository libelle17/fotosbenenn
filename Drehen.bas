Attribute VB_Name = "Drehen"
Dim CommonDialog1 As CommonDialog
Dim cmdSaveAsRotatedjpg As CommandButton
Dim List1 As ListBox
Dim Picture1 As PictureBox
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source!

'----- Anfang Projektdatei GDIPlusRotatejpgLoosless.vbp -----
' Die Komponente 'Microsoft Common Dialog Control 6.0 (comdlg32.ocx)'
' wird benötigt.

'--- Anfang Formular "frmGDIPlusRotatejpgLoosless" alias
' frmGDIPlusRotatejpgLoosless.frm  ---
' Steuerelement: Listen-Steuerelement "List1"
' Steuerelement: Schaltfläche "cmdSaveAsRotatedjpg"
' Steuerelement: Bildfeld-Steuerelement "Picture1"
' Steuerelement: Standarddialog-Steuerelement "CommonDialog1"
' Steuerelement: Schaltfläche "cmdLoadPicture"

Option Explicit

' ----==== GDIPlus Const ====----
Private Const GdiPlusVersion& = 1
Private Const mimejpg As String = "image/jpeg"
Private Const EncoderParameterValueTypeLong& = 4
Private Const EncoderTransformation As String = "{8D0EB2D1-A58E-4EA8-AA14-108074B7B6F9}"

' ----==== Sonstige Types ====----
Private Type PICTDESC
    cbSizeOfStruct As Long
    picType As Long
    hgdiObj As Long
    hPalOrXYExt As Long
End Type

Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7)  As Byte
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

' ----==== GDIPlus Types ====----
Private Type GDIPlusStartupInput
    GdiPlusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type EncoderParameter
    GUID As GUID
    NumberOfValues As Long
    Type As Long
    Value As Long
End Type

Private Type EncoderParameters
    Count As Long
    Parameter(15) As EncoderParameter
End Type

Private Type ImageCodecInfo
    Clsid As GUID
    FormatID As GUID
    CodecNamePtr As Long
    DllNamePtr As Long
    FormatDescriptionPtr As Long
    FilenameExtensionPtr As Long
    MimeTypePtr As Long
    Flags As Long
    Version As Long
    SigCount As Long
    SigSize As Long
    SigPatternPtr As Long
    SigMaskPtr As Long
End Type

' ----==== GDIPlus Enums ====----
Private Enum Status 'GDI+ Status
    OK = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
    ProfileNotFound = 21
End Enum

Private Enum EncoderValueConstants
    EncoderValueColorTypeCMYK = 0
    EncoderValueColorTypeYCCK = 1
    EncoderValueCompressionLZW = 2
    EncoderValueCompressionCCITT3 = 3
    EncoderValueCompressionCCITT4 = 4
    EncoderValueCompressionRle = 5
    EncoderValueCompressionNone = 6
    EncoderValueScanMethodInterlaced = 7
    EncoderValueScanMethodNonInterlaced = 8
    EncoderValueVersionGif87 = 9
    EncoderValueVersionGif89 = 10
    EncoderValueRenderProgressive = 11
    EncoderValueRenderNonProgressive = 12
    EncoderValueTransformRotate90 = 13
    EncoderValueTransformRotate180 = 14
    EncoderValueTransformRotate270 = 15
    EncoderValueTransformFlipHorizontal = 16
    EncoderValueTransformFlipVertical = 17
    EncoderValueMultiFrame = 18
    EncoderValueLastFrame = 19
    EncoderValueFlush = 20
    EncoderValueFrameDimensionTime = 21
    EncoderValueFrameDimensionResolution = 22
    EncoderValueFrameDimensionPage = 23
End Enum

' ----==== Sonstige Enums ====----
Public Enum jpgTransformType
    jpgTransformrotate90 = EncoderValueConstants.EncoderValueTransformRotate90
    jpgTransformRotate180 = EncoderValueConstants.EncoderValueTransformRotate180
    jpgTransformrotate270 = EncoderValueConstants.EncoderValueTransformRotate270
    jpgTransformFlipHorizontal = EncoderValueConstants.EncoderValueTransformFlipHorizontal
    jpgTransformFlipVertical = EncoderValueConstants.EncoderValueTransformFlipVertical
End Enum

' ----==== GDI+ API Declarationen ====----
Private Declare Function GdiplusStartup Lib "gdiplus" _
    (ByRef token As Long, ByRef lpInput As GDIPlusStartupInput, _
    Optional ByRef lpOutput As Any) As Status

Private Declare Function GdiplusShutdown Lib "gdiplus" _
    (ByVal token As Long) As Status

Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" _
    (ByVal FileName As Long, ByRef BITMAP As Long) As Status

Private Declare Function GdipLoadImageFromFile Lib "gdiplus" _
    (ByVal FileName As Long, ByRef image As Long) As Status

Private Declare Function GdipSaveImageToFile Lib "gdiplus" _
    (ByVal image As Long, ByVal FileName As Long, _
    ByRef clsidEncoder As GUID, ByRef encoderParams As Any) As Status

Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" _
    (ByVal BITMAP As Long, ByRef hbmReturn As Long, _
    ByVal background As Long) As Status

Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" _
    (ByRef numEncoders As Long, ByRef Size As Long) As Status

Private Declare Function GdipGetImageEncoders Lib "gdiplus" _
    (ByVal numEncoders As Long, ByVal Size As Long, _
    ByRef Encoders As Any) As Status

Private Declare Function GdipDisposeImage Lib "gdiplus" _
    (ByVal image As Long) As Status

' ----==== OLE API Declarations ====----
Private Declare Function CLSIDFromString Lib "ole32" _
    (ByVal Str As Long, ID As GUID) As Long

Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" _
    (lpPictDesc As PICTDESC, riid As IID, _
    ByVal fOwn As Boolean, lplpvObj As Object)

' ----==== Kernel API Declarations ====----
Private Declare Function lstrlenW Lib "kernel32" _
    (lpString As Any) As Long

Private Declare Function lstrcpyW Lib "kernel32" _
    (lpString1 As Any, lpString2 As Any) As Long

' ----==== Variablen ====----
Private GdipToken As Long
Private GdipInitialized As Boolean
Private InjpgFileName As String
Private InjpgFileTitle&

'------------------------------------------------------
' Funktion     : StartUpGDIPlus
' Beschreibung : Initialisiert GDI+ Instanz
' Übergabewert : GDI+ Version
' Rückgabewert : GDI+ Status
'------------------------------------------------------
Private Function StartUpGDIPlus(ByVal GdipVersion As Long) As Status
    ' Initialisieren der GDI+ Instanz
    Dim GdipStartupInput As GDIPlusStartupInput
    On Error GoTo fehler
    GdipStartupInput.GdiPlusVersion = GdipVersion
    StartUpGDIPlus = GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
    Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in StartUPGDIPlus/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' StartUpGDIPlus

'------------------------------------------------------
' Funktion     : ShutdownGDIPlus
' Beschreibung : Beendet die GDI+ Instanz
' Rückgabewert : GDI+ Status
'------------------------------------------------------
Private Function ShutdownGDIPlus() As Status
    ' Beendet GDI+ Instanz
    On Error GoTo fehler
    ShutdownGDIPlus = GdiplusShutdown(GdipToken)
    Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ShutdownGDIPlus/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' SutdownGDIPlus

'------------------------------------------------------
' Funktion     : Execute
' Beschreibung : Gibt im Fehlerfall die entsprechende
'                GDI+ Fehlermeldung aus
' Übergabewert : GDI+ Status
' Rückgabewert : GDI+ Status
'------------------------------------------------------
Private Function Execute(ByVal lReturn As Status) As Status
    Dim lCurErr As Status
    On Error GoTo fehler
    If lReturn = Status.OK Then
        lCurErr = Status.OK
    Else
        lCurErr = lReturn
        Call MsgBox(GdiErrorString(lReturn) & " GDI+ Error:" & lReturn, vbOKOnly, "GDI Error")
    End If
    Execute = lCurErr
    Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Execute/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Execute

'------------------------------------------------------
' Funktion     : GdiErrorString
' Beschreibung : Umwandlung der GDI+ Statuscodes in Stringcodes
' Übergabewert : GDI+ Status
' Rückgabewert : Fehlercode als String
'------------------------------------------------------
Private Function GdiErrorString$(ByVal lError As Status)
    Dim s$
    On Error GoTo fehler
    Select Case lError
    Case GenericError:              s = "Generic Error."
    Case InvalidParameter:          s = "Invalid Parameter."
    Case OutOfMemory:               s = "Out Of Memory."
    Case ObjectBusy:                s = "Object Busy."
    Case InsufficientBuffer:        s = "Insufficient Buffer."
    Case NotImplemented:            s = "Not Implemented."
    Case Win32Error:                s = "Win32 Error."
    Case WrongState:                s = "Wrong State."
    Case Aborted:                   s = "Aborted."
    Case FileNotFound:              s = "File Not Found."
    Case ValueOverflow:             s = "Value Overflow."
    Case AccessDenied:              s = "Access Denied."
    Case UnknownImageFormat:        s = "Unknown Image Format."
    Case FontFamilyNotFound:        s = "FontFamily Not Found."
    Case FontStyleNotFound:         s = "FontStyle Not Found."
    Case NotTrueTypeFont:           s = "Not TrueType Font."
    Case UnsupportedGdiplusVersion: s = "Unsupported Gdiplus Version."
    Case GdiplusNotInitialized:     s = "Gdiplus Not Initialized."
    Case PropertyNotFound:          s = "Property Not Found."
    Case PropertyNotSupported:      s = "Property Not Supported."
    Case Else:                      s = "Unknown GDI+ Error."
    End Select
    
    GdiErrorString = s
    Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GdiErrorString/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
 End Function ' GdiErrorString

'------------------------------------------------------
' Funktion     : LoadPicturePlus
' Beschreibung : Lädt ein Bilddatei per GDI+
' Übergabewert : Pfad\Dateiname der Bilddatei
' Rückgabewert : StdPicture Objekt
'------------------------------------------------------
Public Function LoadPicturePlus(ByVal FileName As String) As StdPicture
    Dim retStatus As Status
    Dim lBitmap&
    Dim hBitmap&
    On Error GoTo fehler
    ' Öffnet die Bilddatei in lBitmap
    retStatus = Execute(GdipCreateBitmapFromFile(StrPtr(FileName), _
    lBitmap))
    
    If retStatus = OK Then
        
        ' Erzeugen einer GDI Bitmap lBitmap -> hBitmap
        retStatus = Execute(GdipCreateHBITMAPFromBitmap(lBitmap, _
        hBitmap, 0))
        
        If retStatus = OK Then
            ' Erzeugen des StdPicture Objekts von hBitmap
            Set LoadPicturePlus = HandleToPicture(hBitmap, vbPicTypeBitmap)
        End If
        
        ' Lösche lBitmap
        Call Execute(GdipDisposeImage(lBitmap))
        
    End If
    Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in LoadPicturePlus/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LoadPicturePlus

'------------------------------------------------------
' Funktion     : RotatejpgLossless
' Beschreibung : Verlustfreies Rotieren von jpg´s per GDI+
' Übergabewert : InFileName = Pfad\Dateiname.jpg
'                OutFileName = Pfad\Dateiname.jpg
'                jpgTransform = Rotationstyp
' Rückgabewert : True = rotation erfolgreich
'                False = rotation fehlgeschlagen
'------------------------------------------------------
Function RotatejpgLossless(ByVal InFilename As String, _
    ByVal OutFilename As String, _
    ByVal jpgTransform As jpgTransformType) As Boolean
    
    Dim retStatus&
    Dim lBitmap&
    Dim RetVal As Boolean
    On Error GoTo fehler
    retStatus = Execute(GdipLoadImageFromFile(StrPtr(InFilename), lBitmap))
    
    If retStatus = OK Then
        
        Dim PicEncoder As GUID
        Dim tParams As EncoderParameters
        
        '// Ermitteln der CLSID vom mimeType Encoder
        RetVal = GetEncoderClsid(mimejpg, PicEncoder)
        If RetVal = True Then
            
            ' Initialisieren der Encoderparameter
            tParams.Count = 1
            With tParams.Parameter(0) ' Transformation
                ' Setzen der Transformations GUID
                CLSIDFromString StrPtr(EncoderTransformation), .GUID
                .NumberOfValues = 1
                .Type = EncoderParameterValueTypeLong
                .Value = VarPtr(jpgTransform)
            End With
            
            ' Speichert lBitmap als jpg
            retStatus = Execute(GdipSaveImageToFile(lBitmap, _
                    StrPtr(OutFilename), PicEncoder, tParams))
            
            If retStatus = OK Then
                RotatejpgLossless = True
            Else
                RotatejpgLossless = False
            End If
            
        Else
            RotatejpgLossless = False
            MsgBox "Konnte keinen passenden Encoder ermitteln.", _
            vbOKOnly, "Encoder Error"
        End If
        
        ' Lösche lBitmap
        Call Execute(GdipDisposeImage(lBitmap))
    End If
    Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in RotatejpgLossless/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' RotatejpgLossless

'------------------------------------------------------
' Funktion     : HandleToPicture
' Beschreibung : Umwandeln einer GDI+ Bitmap Handle in ein StdPicture
' Objekt
' Übergabewert : hGDIHandle = GDI+ Bitmap Handle
'                ObjectType = Bitmaptyp
' Rückgabewert : StdPicture Objekt
'------------------------------------------------------
Private Function HandleToPicture(ByVal hGDIHandle&, ByVal ObjectType As PictureTypeConstants, _
    Optional ByVal hpal& = 0) As StdPicture
    
    Dim tPictDesc As PICTDESC
    Dim IID_IPicture As IID
    Dim oPicture As IPicture
    On Error GoTo fehler
    ' Initialisiert die PICTDESC Structur
    With tPictDesc
        .cbSizeOfStruct = Len(tPictDesc)
        .picType = ObjectType
        .hgdiObj = hGDIHandle
        .hPalOrXYExt = hpal
    End With
    
    ' Initialisiert das IPicture Interface ID
    With IID_IPicture
        .Data1 = &H7BF80981
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(3) = &HAA
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    
    ' Erzeugen des Objekts
    OleCreatePictureIndirect tPictDesc, IID_IPicture, True, oPicture
    
    ' Rückgabe des Pictureobjekts
    Set HandleToPicture = oPicture
    Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in HandelToPicture/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' HandleToPicture

'------------------------------------------------------
' Funktion     : GetEncoderClsid
' Beschreibung : Ermittelt die Clsid des Encoders
' Übergabewert : mimeType = mimeType des Encoders
'                pClsid = CLSID des Encoders (in/out)
' Rückgabewert : True = Ermitteln erfolgreich
'                False = Ermitteln fehlgeschlagen
'------------------------------------------------------
Private Function GetEncoderClsid%(mimeType$, pClsid As GUID)
    
    Dim num&
    Dim Size&
    Dim pImageCodecInfo() As ImageCodecInfo
    Dim j&
    Dim buffer$
    On Error GoTo fehler
    Call GdipGetImageEncodersSize(num, Size)
    If (Size = 0) Then
        GetEncoderClsid = False  '// fehlgeschlagen
        Exit Function
    End If
    
    ReDim pImageCodecInfo(0 To Size \ Len(pImageCodecInfo(0)) - 1)
    Call GdipGetImageEncoders(num, Size, pImageCodecInfo(0))
    
    For j = 0 To num - 1
        buffer = Space$(lstrlenW(ByVal pImageCodecInfo(j).MimeTypePtr))
        
        Call lstrcpyW(ByVal StrPtr(buffer), _
            ByVal pImageCodecInfo(j).MimeTypePtr)
        
        If (StrComp(buffer, mimeType, vbTextCompare) = 0) Then
            pClsid = pImageCodecInfo(j).Clsid
            Erase pImageCodecInfo
            GetEncoderClsid = True  '// erfolgreich
            Exit Function
        End If
    Next j
    
    Erase pImageCodecInfo
    GetEncoderClsid = False  '// fehlgeschlagen
    Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetEncoderClsid/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GetEncoderClsid
Private Sub cmdLoadPicture_Click()
    On Error GoTo fehler
    
    If GdipInitialized = True Then
        
        With CommonDialog1
            .Filter = "jpg Files (*.jpg)|*.jpg"
            .CancelError = True
            .ShowOpen
        End With
        
        Picture1.Picture = LoadPicturePlus(CommonDialog1.FileName)
        InjpgFileName = CommonDialog1.FileName
        InjpgFileTitle = CommonDialog1.FileTitle
        
        If Not Picture1.Picture Is Nothing Then cmdSaveAsRotatedjpg.Enabled = True
    End If
    
    Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in cmdLoadPicture_Click/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' cmdLoadPicture_Click

Private Sub cmdSaveAsRotatedjpg_Click(FDt As FDatei)
    If GdipInitialized = True Then
        Dim RetVal As Boolean
        Dim Transform As jpgTransformType
        Dim TransformFileName As String
        
        Select Case List1.List(List1.ListIndex)
        Case "jpgTransformRotate90"
            Transform = jpgTransformrotate90
        Case "jpgTransformRotate180"
            Transform = jpgTransformRotate180
        Case "jpgTransformRotate270"
            Transform = jpgTransformrotate270
        Case "jpgTransformFlipHorizontal"
            Transform = jpgTransformFlipHorizontal
        Case "jpgTransformFlipVertical"
            Transform = jpgTransformFlipVertical
        End Select
        
        'temporären Dateinamen erstellen
        TransformFileName = Mid$(InjpgFileName, 1, Len(InjpgFileName) - Len(InjpgFileTitle)) & "_" & InjpgFileTitle
        
        ' jpg transformieren
        RetVal = RotatejpgLossless(InjpgFileName, TransformFileName, Transform)
        
        If RetVal = True Then
            'lösche Originaldatei
'            Kill InjpgFileName
            Call FDt.LöscheDatei(InjpgFileName)
            
            'temporäre Datei in Original umbenennen
'            Name TransformFileName As InjpgFileName
            Call FDt.VerschiebeFD(TransformFileName, InjpgFileName)
            
            'Datei wieder laden und anzeigen
            Picture1.Picture = LoadPicturePlus(InjpgFileName)
        Else
            MsgBox "Das rotieren der jpg ist fehlgeschlagen.", _
                vbOKOnly, "Error"
        End If
        
    End If
    Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in cmdSaveAsRotatedjpg_Click/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' cmdSaveAsRotatedjpg_Click

Sub DForm_Load()
    Dim retStatus As Status
    GdipInitialized = False
    On Error GoTo fehler
    retStatus = Execute(StartUpGDIPlus(GdiPlusVersion))
    If retStatus = OK Then
        GdipInitialized = True
    Else
        MsgBox "GDI+ not inizialized.", vbOKOnly, "GDI Error"
    End If
    On Error Resume Next
    cmdSaveAsRotatedjpg.Enabled = False
    List1.AddItem "jpgTransformRotate90"
    List1.AddItem "jpgTransformRotate180"
    List1.AddItem "jpgTransformRotate270"
    List1.AddItem "jpgTransformFlipHorizontal"
    List1.AddItem "jpgTransformFlipVertical"
    List1.ListIndex = 0
    Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in DForm_Load/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' DForm_Load

Sub DForm_Unload(Cancel As Integer)
    Dim retStatus As Status
    On Error GoTo fehler
    If GdipInitialized = True Then
        retStatus = Execute(ShutdownGDIPlus)
    End If
    Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in DForm_Unlaod/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' DForm_Unload
'--- Ende Formular "frmGDIPlusRotatejpgLoosless" alias
' frmGDIPlusRotatejpgLoosless.frm  ---
'------ Ende Projektdatei GDIPlusRotatejpgLoosless.vbp ------

