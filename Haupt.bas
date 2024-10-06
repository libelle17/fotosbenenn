Attribute VB_Name = "Haupt"
Option Explicit
Public Const vNS$ = vbNullString
Public DBCn As New Connection
Public obTrans%

Declare Function GetLogicalDriveStrings& Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength&, ByVal lpBuffer$)
Declare Function sndPlaySound32& Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName$, ByVal uFlags&)
Declare Function GetDriveType& Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive$)
Public IViewPfad$
Const IniDat$ = "FotosBenenn.ini"
Public Const uVerz$ = "u:\"
Public FPos&, FNr& ' Fehlerposition
'Public wsh As IWshShell_Class ' %windir%\system32\wshom.ocx
'Public WMIreg As SWbemObjectEx ' %windir%\system32\wbem\wbemdisp.tlb

'Public ArchPat$, Quelle$, Archiv$, SteuerDB$ ', obAlteArch%, AlteArch$, obQuelle%, obEingel%, Eingel$
'Dim Verz(), Datei()
Const REG_DWORD = 4
Const REG_BINARY = 3

'Public QDB As DAO.Database ' NamDB, u:\anamnese\quelle.mdb
'Public SDB As DAO.Database ' SteuerDB, u:\fotosinp.mdb
'Public rBld As DAO.Recordset

'Dim nz$
Public FSO As FileSystemObject
Public Const RegStelle$ = "Software\GSProducts\FotosBenenn"
'Dim sC(1) As Collection, sCa&(1), sCe&(1) ' 0 = Archiv, 1 = Eingelesene
'Public Dateien() As File, indDat&
Public GrößenVerhältnis!
'Public rSteu As DAO.Recordset
Public rSteu As New ADODB.Recordset
Public Pat_id&
Public Const sIZahl = 2
Public frMü As fürIcon
'Public Constr$ ' Connection String
Public QuelCStr$ ' Connection String zur Patientendatenbank
Public FotoCStr$ ' Connection String zur FotoDatenbank
Public FotoCn As New ADODB.Connection
Public obMySQL%
Public Const LiName = "linux1", LiServer$ = "\\" & LiName & "\" ' \\linux1\
Public pVerz$


' wird nirgends anders aufgerufen
Sub Main()
' nz = Chr(13)
 On Error GoTo fehler
 pVerz = IIf(Dir("p:") <> "", "p:", LiServer & "Daten\Patientendokumente") & "\"
 Set frMü = New fürIcon
 frMü.Show
 Exit Sub
' Call mainaltrest

 Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in main/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' main
' kommt nur in DateiAnzeig vor
Function GrVerhErm(frm As fürIcon, Datei$)
 Dim strDummy As String
 Dim ff As Integer
 Dim c As Integer
 Dim s As String
 Dim L As Long
 Dim jpgWidth As Long
 Dim jpgHeight As Long
 On Error GoTo fehler
 ff = FreeFile()
 Open Datei For Binary Access Read As #ff
  ' Test auf JPEG-Datei
 If Input(2, #ff) <> (Chr$(&HFF) & Chr$(&HD8)) Then
  Close #ff
  Exit Function
 End If
 strDummy = Input(2, #ff)
 Do
  L = Asc(Input(1, #ff))
  L = L * 256 + Asc(Input(1, #ff))
  s = Input(L - 2, #ff)
  If c = &HC0 Or c = &HC2 Then
   jpgWidth = Asc(Mid$(s, 4, 1))
   jpgWidth = jpgWidth * 256 + Asc(Mid$(s, 5, 1))
   jpgHeight = Asc(Mid$(s, 2, 1))
   jpgHeight = jpgHeight * 256 + Asc(Mid$(s, 3, 1))
  End If
  If Input(1, #ff) <> Chr$(255) Then
   Exit Do
  End If
  c = Asc(Input(1, #ff))
  Loop While c <> &HD9
   Close #ff
  ' Anzeige der ermittelten Information:
  frm.DateiHöhe = jpgHeight
  frm.DateiBreite = jpgWidth
  If jpgHeight <> 0 Then GrößenVerhältnis = jpgWidth / jpgHeight
'  MsgBox "Die Grafik in test.jpg ist " & vbNewLine & _
'   CStr(jpgWidth) & " Pixel breit und " & vbNewLine & _
'   CStr(jpgHeight) & " Pixel hoch.", _
'   vbInformation, "jpg-Analyse"
  Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GrVerhErm/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GrVerhErm

Public Function Sound(Pfad$)
' Static altPfad$
' If Pfad <> vns Then
'  altPfad = Pfad
' End If
' Call sndPlaySound32(altPfad, 1)
 Call sndPlaySound32(Pfad, 1)
End Function ' Sound


Sub fülleLw(frm)
 Dim drv As Drive, obIM%, drvVolName$, FNr$, FotoV1$, FotoV2$
 Dim Fld As Folder, Fil As File
 On Error GoTo fehler
 For Each drv In FSO.Drives
  If drv.DriveType = Removable Then
    On Error Resume Next
    drvVolName = drv.VolumeName
    FNr = Err.Number
    On Error GoTo fehler
    If FNr > 0 Then drvVolName = Err.Description
    Call frm.Lw.AddItem(drv.Path + " " + drvVolName)
    If FNr = 0 Then
'     If drv.TotalSize > 1457664 Then
      obIM = 0
      FotoV1 = drv.RootFolder.Path + "DCIM"
      If FSO.FolderExists(FotoV1) Then
       For Each Fld In FSO.GetFolder(FotoV1).SubFolders ' 28.9.08 wg. Canon
'        FotoV2 = FotoV1 + "\100DSCIM"
'       If FSO.FolderExists(FotoV2) Then
'        If FSO.GetFolder(FotoV2).Files.Count > 0 Then
        For Each Fil In Fld.Files
         If Not Fil Like "*.modd" Then
          If DateiArt(Fil.name) <> 0 Then
           obIM = -1
           Exit For
          End If
         End If
        Next Fil
        If obIM Then Exit For
       Next Fld
      End If
      If obIM Then
       frm.Lw.ListIndex = frm.Lw.ListCount - 1
      End If
'     End If
    End If
  End If
 Next drv
 Call frm.Lw.AddItem("(keine Datenübertragung vom Foto)")
 If frm.Lw.ListIndex = -1 Then
  frm.Lw.ListIndex = frm.Lw.ListCount - 1
 End If
Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in fülleLw/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' fülleLW

' kommt vor in doPatNameChange, SchreibDatensatz und FDC_indvorWechsel
Function getPat_id&(PatName$)
 Dim Spl$()
 If IsNull(PatName) Or PatName = vNS Then
  getPat_id = -1
 Else
  Spl = Split(PatName, "|")
  If UBound(Spl) < 1 Then
   getPat_id = -1
  Else
   If Not IsNumeric(Spl(1)) Then
    getPat_id = -1
   Else
    getPat_id = CLng(Spl(1))
   End If
  End If
 End If
End Function ' getPat_id

Function doAlleÜbertragen(frm As fürIcon)
 Dim i&, erg&
 On Error GoTo fehler
' Call frm.FDC(frm.FDC.indDat).aktSpeichern(frm)
 If Not IsNumeric(frm.Kompressionsgrad) Then
  MsgBox "Der Kompressionsgrad ist nicht numerisch. Breche ab!"
  Exit Function
 End If
 If frm.Kompressionsgrad > 100 Then
  MsgBox "Der Kompressionsgrad ist größer 100. Breche ab!"
  Exit Function
 End If
 If frm.Kompressionsgrad < 10 Then
  MsgBox "Der Kompressionsgrad ist kleiner 10. Breche ab!"
  Exit Function
 End If
 erg = MsgBox("Wollen Sie wirklich alle fertigen Dateien mit dem Kompressionsgrad " & frm.Kompressionsgrad & " übertragen", vbYesNo)
 If erg = vbYes Then
'  Call frm.FDC(frm.FDC.indDat).aktSpeichern(frm)
  frm.FDC.indDat = frm.FDC.indDat
  For i = 1 To frm.FDC.Count
   If FSO.FileExists(frm.FDC(i).bm) Then
    Call doAusgabe("Abspeichern von Bild " & i & "/" & frm.FDC.Count & ": " & frm.FDC(i).Fil.Path & " - > " & frm.Kompr & frm.FDC(i).NeuerName)
    Call frm.FDC(i).doÜbertragen(frm.Kompr, frm.ArchPat, frm.Archiv, frm.HScroll1, frm.HScroll2, frm.HScroll3, frm.Kompressionsgrad, frm.stehenLassen)
   End If
weiter:
  Next i
  Call do_Start(frm)
 End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doAlleÜbertragen/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doAlleÜbertragen

Public Function getVariantePfad$(Fil As File)
' On Error Resume Next
 getVariantePfad = Left(Fil.Path, Len(Fil.Path) - Len(Fil.name)) & "v" & Fil.name
' On Error GoTo 0
End Function

Public Function datForm(dat) ' for vb-Datumsformat oder vb-double (#)
 On Error GoTo fehler
 If IsNull(dat) Then
  datForm = "null"
 ElseIf obMySQL Then
  datForm = "'" + Format(dat, "yyyy-mm-dd hh:mm:ss") + "'"
 Else
  datForm = "#" + Format(dat, "mm\/dd\/yy hh:mm:ss") + "#"
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in datForm/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' datForm

Public Function DateiArt%(ByVal DName$, Optional ByRef Nr&) ' 1 = Bild, 2 = Ton
' da like "PICT####" sehr viel Zeit verbrauchen soll
 Dim pos&, pose&
 DName = UCase$(DName)
 Do While True
  pos = InStr(DName, "\")
  If pos = 0 Then Exit Do
  DName = Mid(DName, pos + 1)
 Loop
 If Not DName Like "V*" Then
  If InStrB(DName, "JPG") <> 0 Then
   pos = InStr(DName, "BILD")
   If pos > 0 Then
    DateiArt = 1
   Else
    pos = InStr(DName, "PICT") ' "PICT" 5.2.14 Olympus-Foto
    If pos > 0 Then
     DateiArt = 1
    Else
     pos = InStr(DName, "P") ' "P" 28.6.14 Olympus-Foto
     If pos = 1 And IsNumeric(Mid(DName, 3, 6)) Then
      DateiArt = 3
      Nr = CLng(Mid(DName, 3, 6))
     Else
      pos = InStr(DName, "IMG_")
      If pos > 0 Then
       DateiArt = 1
      Else
       pos = InStr(DName, "SAM_")
       If pos > 0 Then
        DateiArt = 1
       Else
        pos = InStr(DName, "DSCN")
        If pos > 0 Then
         DateiArt = 1
        Else
         pos = InStr(DName, "201")
         If pos > 0 Then
          DateiArt = 4
         Else
          pos = InStr(DName, "CIMG")
          If pos > 0 Then
           DateiArt = 1
          Else
           pos = InStr(DName, "DSC_")
          If pos > 0 Then
           DateiArt = 1
          Else
           pos = InStr(DName, "DSC")
           If pos > 0 Then
            DateiArt = 5
           Else
'            MsgBox DName + " nicht klassifiziert"
'           Stop
            End If
           End If
          End If
         End If
        End If
       End If
      End If
     End If
    End If
   End If
  ElseIf InStrB(DName, "WAV") <> 0 Then ' instrb(dname, "JPG"
   pos = InStr(DName, "BILD")
   If pos > 0 Then
    DateiArt = 2
   Else
    pos = InStr(DName, "P") ' "PICT" 5.2.14 Olympus-Foto
    If pos > 0 Then
     DateiArt = 2
    Else
     pos = InStr(DName, "SND_")
     If pos = 0 Then pos = InStr(DName, "IMG_")
     If pos > 0 Then
      DateiArt = 2
     Else
     End If
    End If
   End If
  End If
 End If
 If DateiArt = 5 Then
  If IsNumeric(Mid(DName, pos + 3, 5)) Then
   Nr = CLng(Mid(DName, pos + 3, 5))
   DateiArt = 1
  Else
   DateiArt = 0
  End If
 ElseIf DateiArt = 1 Or DateiArt = 2 Then
  If IsNumeric(Mid(DName, pos + 4, 4)) Then
   Nr = CLng(Mid(DName, pos + 4, 4))
  Else
   DateiArt = 0
  End If
 ElseIf DateiArt = 4 Then
  If IsNumeric(Mid(DName, pos + 9, 6)) Then
   Nr = CLng(Mid(DName, pos + 9, 6))
  Else
   DateiArt = 0
  End If
  DateiArt = 1
 ElseIf DateiArt = 3 Then
  DateiArt = 1
 End If
End Function ' Dateiart

Public Function zurück()
 Dim erg$
 On Error Resume Next
 erg = Dir("p:\fotos neu\*.*")
 Do While erg <> vNS
  If Len(erg) > 10 Then
   Name "p:\fotos neu\" & erg As "p:\fotos neu\" & Left(erg, 8) & Right(erg, 4)
  End If
  erg = Dir
 Loop
End Function ' zurück

Public Function ProgEnde()
 End
End Function

