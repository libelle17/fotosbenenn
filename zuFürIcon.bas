Attribute VB_Name = "zuF�rIcon"
Option Explicit
Const sQuelle$ = "P:\Fotos neu\"
Const sSteuerDB$ = "U:\fotosinp.mdb"
Const sNamDB$ = "U:\Anamnese\Quelle.mdb"
Const sKompr$ = "P:\"
Const sArchPat$ = "P:\Fotos alt"
Const sArchiv$ = "T:\Fotos"
Const sKompressionsgrad$ = "30"
Public imAufbau As Boolean
Public Const CStrAcc$ = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
Dim ErrDescription$
'Public Const CStrMy$ = "DRIVER={MySQL ODBC 3.51 Driver};server=linux1;user=praxis;pwd=sonne;database="

'Const sobAltearch% = 1
'Const sAlteArch$ = sArchPat + "HDI alt"
'Const sobQuelle% = 1
'Const sobEingel% = 0
'Const SEingel$ = sArchPat + "eingelesen\"
' wird nur in do_Start aufgerufen
Function DatKop(frm As f�rIcon, Dv$)
 Dim Qfol As Folder
 On Error GoTo fehler
 Set Qfol = FSO.GetFolder(frm.Quelle) ' "P:\Fotos neu"
 Dim ZFolTmp As Folder
 Dim drv As Drive
 Set drv = FSO.GetDrive(Dv)
 Dim ZFoltmpStr$
 ZFoltmpStr = frm.Quelle
 If Right(ZFoltmpStr, 1) = "\" Then
  ZFoltmpStr = Left(ZFoltmpStr, Len(ZFoltmpStr) - 1)
 End If
 ZFoltmpStr = ZFoltmpStr + " tmp"
 If FSO.FolderExists(ZFoltmpStr) Then
  Set ZFolTmp = FSO.GetFolder(ZFoltmpStr)
  If ZFolTmp.Files.Count > 0 Or ZFolTmp.SubFolders.Count > 0 Then
   MsgBox "Ordner " + ZFolTmp.Path + " nicht leer. Bitte leeren, dann nochmal aufrufen!"
   Unload frm
  End If
 Else
  Set ZFolTmp = ErstelleOrdner(ZFoltmpStr, frm)
 End If
 Call doBewegInRoot(drv.RootFolder, ZFolTmp, frm) ' z.B. G: P:\HDI neu tmp
' Call QuelleArchivieren(Qfol)
 Call doBewegInRoot(ZFolTmp, Qfol, frm)
 Call L�scheOrdner(ZFolTmp.Path, frm)
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in DatKop/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DatKop

' wird nur in DatKop aufgerufen
Function doBewegInRoot(Quelle As Folder, ArchPat As Folder, frm As f�rIcon) ' mit "\"
 Dim ZFol As Folder, Fol As Folder, Fil As File
 Dim ZFolStr$
 Dim erg&, DArt%
 On Error GoTo fehler
 For Each Fol In Quelle.SubFolders
  ZFolStr = ArchPat.Path '+ "\" + Fol.Name
  If FSO.FolderExists(ZFolStr) Then
   Set ZFol = FSO.GetFolder(ZFolStr)
  Else
   Set ZFol = ErstelleOrdner(ZFolStr, frm)
  End If
  Call doBewegInRoot(Fol, ZFol, frm)
'  Call FSO.DeleteFolder(Fol.Path, False)
 Next Fol
 Call doAusgabe("doBewegInRoot: " & Quelle.Files.Count & " Dateien in " & Quelle)
 For Each Fil In Quelle.Files
  If Not Fil Like "*.modd" Then
   DArt = DateiArt(Fil.name)
   If DArt <> 0 Then '  If Fil.Name Like "PICT*" Or Fil.Name Like "BILD*" Then
'   If LCase(Fil.Path) Like "*.wav" Then Stop
    Call VerschiebeFI(Fil.Path, ArchPat.Path & IIf(Right(ArchPat.Path, 1) = "\", vNS, "\") & Fil.name, frm, DArt)
    DoEvents
   Else
    Call doAusgabe(Fil & " erf�llt nicht die Kriterien ")
   End If
  End If
 Next Fil
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doBeweginRoot/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doBeweginRoot

' kommt vor in DoBeweg, Datkop und DoBeweginRoot
Function ErstelleOrdner(Vol$, frm As f�rIcon, Optional unsicher%) As Folder
 Dim Ausgabe$
 On Error GoTo fehler
 If unsicher Then
  Ausgabe = "Versuche zu Erstellen: "
  On Error Resume Next
 Else
  Ausgabe = "Erstelle: "
 End If
 Call doAusgabe(Ausgabe & Vol)
 Set ErstelleOrdner = FSO.CreateFolder(Vol)
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ErstelleOrdner/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ErstelleOrdner

' kommt vor in doBeweg und DatKop
Function L�scheOrdner(Vol$, frm As f�rIcon, Optional unsicher%)
 Dim Ausgabe$
 On Error GoTo fehler
 If unsicher Then
  Ausgabe = "Versuche zu Entfernen: "
  On Error Resume Next
 Else
  Ausgabe = "Entferne: "
 End If
 Call doAusgabe(Ausgabe & Vol)
 Call FSO.DeleteFolder(Vol)
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in L�scheOrdner/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' L�scheOrdner

' kommt vor in Rotier, cmdSaveAsRotated_jpg
Function L�scheDatei(D1$, frm As f�rIcon, Optional unsicher%)
 Dim Ausgabe$
 On Error GoTo fehler
 If unsicher Then
  Ausgabe = "Versuche zu L�schen: "
  On Error Resume Next
 Else
  Ausgabe = "L�sche: "
 End If
 Call doAusgabe(Ausgabe & D1)
 Call FSO.DeleteFile(D1)
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in L�scheDatei/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' L�scheDatei

Function doEnable(frm As f�rIcon, Status%)
 Dim e1%, E2%
 On Error GoTo fehler
 Select Case Status
  Case 1: e1 = True: E2 = False
  Case 2: e1 = False: E2 = True
 End Select
 frm.Lw.Enabled = e1
 frm.Start.Enabled = e1
 frm.Quelle.Enabled = e1
 frm.QuelleLabel.Enabled = e1
 frm.SteuerDB.Enabled = e1
 frm.SteuerDBBez.Enabled = e1
 frm.NamDB.Enabled = e1
 frm.NamDBBez.Enabled = e1
 frm.Kompr.Enabled = e1
 frm.KomprBez.Enabled = e1
 frm.ArchPat.Enabled = e1
 frm.ArchPatBez.Enabled = e1
 frm.Archiv.Enabled = e1
 frm.ArchivBez.Enabled = e1
 frm.Schrittweite.Enabled = E2
 frm.SchrittweiteBez.Enabled = E2
 frm.ZumAnfang.Enabled = E2
 frm.R�ckw�rtsCmd.Enabled = E2
 frm.Vorw�rtsCmd.Enabled = E2
 frm.ldP.Enabled = E2
 frm.ndP.Enabled = E2
 frm.ZumEnde.Enabled = E2
 frm.ersterOffenerCmd.Enabled = E2
 frm.Lad.Enabled = E2
 frm.TonCmd.Enabled = E2
 frm.PatName.Enabled = E2
 frm.PatNamBez.Enabled = E2
 frm.obPat.Enabled = E2
 frm.wieVoriges.Enabled = E2
 frm.wieN�chstes.Enabled = E2
 frm.wieLetztesdPat.Enabled = E2
 frm.n�dP.Enabled = E2
 frm.K�rperTeil.Enabled = E2
 frm.K�rperTeilBez.Enabled = E2
 frm.BildDatum.Enabled = E2
 frm.BildDatumBez.Enabled = E2
 frm.FarbenZur�ck.Enabled = E2
 frm.HellLab.Enabled = E2
 frm.keinTon.Enabled = E2
 frm.stehenLassen.Enabled = E2
 frm.Kompressionsgrad.Enabled = E2
 frm.KompressionsgradLbl.Enabled = E2
 frm.Beschreibung.Enabled = E2
 frm.BeschreibungBez.Enabled = E2
 frm.Wagner.Enabled = E2
 'frm.WagnerL.Enabled = E2
 frm.WagnerBez.Enabled = E2
 frm.Armstrong.Enabled = E2
 frm.ArmstrongBez.Enabled = E2
 frm.Dopplerlabel.Enabled = E2
 frm.Fu�statusBez.Enabled = E2
 frm.LinksCmd.Enabled = E2
 frm.rechtsCmd.Enabled = E2
 frm.FreiDrehen.Enabled = False
 frm.LinksNeu.Enabled = False
 frm.RechtsNeu.Enabled = False
 frm.nur2.Enabled = E2
 frm.FarbenZur�ck = E2
 frm.Alle�bertragen.Enabled = E2
 frm.rueckgaengig.Enabled = Not E2
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doEnable/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doEnable

' kommt als CommandButton vor
Function Rotier(Richtung, frm As f�rIcon)
 Dim anf$, tmp$, FTe As FileTimeEnum, DDat#
 On Error GoTo fehler
 Call frm.BeginWarten
 anf = frm.FDC(frm.FDC.indDat).Fil.Path
 tmp = Left(frm.FDC(frm.FDC.indDat).Fil.Path, Len(frm.FDC(frm.FDC.indDat).Fil.Path) - 4) + "_.jpg"
 If RotatejpgLossless(anf, tmp, Richtung) = -1 Then
  For FTe = 1 To 4
   If FTe <> 2 Then
    DDat = GetFileTime(anf, FTe)
    Call SetFileTimeByDate(tmp, FTe, DDat)
   End If
  Next FTe
'  Kill anf
  Call L�scheDatei(anf, frm)
'  Name tmp As anf
  Call VerschiebeFI(tmp, anf, frm)
  Call DateiAnzeig(frm)
 End If
 Call frm.EndeWarten
 Exit Function ' rotier
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in rotier/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' exit function

Function doAusgabe(Str$, Optional frm As f�rIcon)
 Dim ifrm As Form
 If frm Is Nothing Then
  For Each ifrm In Forms
   If ifrm.name = "f�rIcon" Then
    Set frm = ifrm
    Exit For
   End If
  Next ifrm
 End If
 If frm Is Nothing Then
  Set frm = f�rIcon
 End If
 If frm.Ausgabe = vNS Then frm.Ausgabe = Str Else frm.Ausgabe = frm.Ausgabe & vbCrLf & Str
End Function ' doAusgabe(Str$, frm As f�rIcon)

Function VerschiebeFI(D1$, D2$, frm As f�rIcon, Optional DArt% = -1, Optional unsicher%)
 Dim Ausgabe$
 On Error GoTo fehler
 If unsicher Then
  Ausgabe = "Versuche zu Verschieben: "
  On Error Resume Next
 Else
  Ausgabe = "Verschiebe: "
 End If
 
 Dim D2a$, pos& ' 25.5.08
 D2a = D2
 If DArt = -1 Then DArt = DateiArt(D2a)
 If DArt <> 0 Then 'UCase(D2a) Like "*BILD####*" Or UCase(D2a) Like "*PICT####*" Then
  If Not D2a Like "*######## ######*" Then
   If D2a Like "*.???" Then
    D2a = Left(D2a, Len(D2a) - 4) & Format(FileDateTime(D1), " yyyymmdd hhmmss") & Right(D2a, 4)
   ElseIf D2a Like "*.??" Then
    D2a = Left(D2a, Len(D2a) - 3) & Format(FileDateTime(D1), " yyyymmdd hhmmss") & Right(D2a, 3)
   ElseIf D2a Like "*.?" Then
    D2a = Left(D2a, Len(D2a) - 2) & Format(FileDateTime(D1), " yyyymmdd hhmmss") & Right(D2a, 2)
   ElseIf D2a Like "*." Then
    D2a = Left(D2a, Len(D2a) - 1) & Format(FileDateTime(D1), " yyyymmdd hhmmss") & Right(D2a, 1)
   Else
    D2a = Left(D2a, Len(D2a) - 0) & Format(FileDateTime(D1), " yyyymmdd hhmmss")
   End If
  End If
 End If
 
 Call doAusgabe(Ausgabe & D1 & " -> " & D2a)
 If FSO.FileExists(D2a) Then
  Call FSO.MoveFile(D2a, REPLACE(REPLACE(LCase(D2a), ".jpg", " vorher.jpg"), ".wav", " vorher.wav"))
 End If
 Call FSO.MoveFile(D1, REPLACE(D2a, "SND_", "IMG_"))
 VerschiebeFI = D2a
 Exit Function
fehler:
ErrDescription = Err.Description
ErrNumber = Err.Number
If InStrB(ErrDescription, "existiert bereits") > 0 Then
 Kill REPLACE(REPLACE(LCase(D2a), ".jpg", " vorher.jpg"), ".wav", " vorher.wav")
 Resume
End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in VerschiebeFI/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' VerschiebeFI

' kommt vor in DateiLad, Rotier
Function DateiAnzeig(frm As f�rIcon)
 Dim Bildlad%
 On Error GoTo fehler
 Call GrVerhErm(frm, frm.FDC(frm.FDC.indDat).Fil.Path)
 frm.DateiPfad = frm.FDC(frm.FDC.indDat).Fil.Path
 frm.DateiZeit = FSO.GetFile(frm.FDC(frm.FDC.indDat).Fil.Path).DateLastModified
 frm.BildDatum.Value = frm.DateiZeit 'DateValue(frm.DateiZeit)
 If frm.BildDatum.Value <> frm.DateiZeit Then
  frm.BildDatum.Value = frm.DateiZeit
 End If
 frm.Position = frm.FDC.indDat
 Call frm.Form_Resize
 Dim Datei$, obabweich%, obabwDatei%
 obabweich = 0
 obabwDatei = 0
 If frm.HScroll1 <> 90 Or frm.HScroll2 <> 90 Or frm.HScroll3 <> 250 Then
  obabweich = True
  Datei = frm.Quelle & "v" & frm.FDC(frm.FDC.indDat).Fil.name
  If FSO.FileExists(Datei) Then obabwDatei = True
 End If
 If Not obabweich Or Not obabwDatei Then Datei = frm.FDC(frm.FDC.indDat).Fil.Path
 Bildlad = -1
 frm.Image1.Picture = LoadPicture(Datei)
 Bildlad = -2
 frm.Picture2.Picture = LoadPicture(Datei)
 Bildlad = 0
 If obabweich And Not obabwDatei Then Call f�rIcon.HellKontr
 Exit Function
fehler:
If Bildlad = -1 Then
 frm.Image1.Picture = LoadPicture(App.Path + "\..\icons\Mug of Tea.ico")
 Resume Next
ElseIf Bildlad = -2 Then
 frm.Picture2.Picture = LoadPicture(App.Path + "\..\icons\Mug of Tea.ico")
 Resume Next
End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in DateiAnzeig/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DateiAnzeig

Function doFindeN�chstendPat(frm As f�rIcon)
 Dim i&, obgefunden%
 On Error GoTo fehler
 If frm.PatName <> vNS Then
  obgefunden = 0
  For i = frm.FDC.indDat + 1 To frm.FDC.Count
   If frm.FDC(i).PatName = frm.PatName Then
    obgefunden = True
    Exit For
   End If
  Next i
 End If
 If obgefunden Then
  frm.FDC.indDat = i
 End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doFindeN�chstendPat/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doFindeN�chstendPat
Function doFindeLetztendPat(frm As f�rIcon)
 Dim i&, obgefunden%
 On Error GoTo fehler
 If frm.PatName <> vNS Then
  obgefunden = 0
  For i = frm.FDC.indDat - 1 To 1 Step -1
   If frm.FDC(i).PatName = frm.PatName Then
    obgefunden = True
    Exit For
   End If
  Next i
 End If
 If obgefunden Then
  frm.FDC.indDat = i
 End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doFindeLetztendPat/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doFindeLetztendPat

Function doVorw�rtsCmd(frm As f�rIcon)
 On Error GoTo fehler
 If Not IsNumeric(frm.Schrittweite) Then frm.Schrittweite = 1
 frm.FDC.indDat = frm.FDC.indDat + frm.Schrittweite
 If frm.FDC.indDat > frm.FDC.Count Then frm.FDC.indDat = frm.FDC.Count
 ' Call frm.FDC(frm.FDC.indDat).DateiLad(frm)
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doVorw�rtsCmd/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doVorw�rtsCmd

Function doR�ckw�rtsCmd(frm As f�rIcon)
 On Error GoTo fehler
 If Not IsNumeric(frm.Schrittweite) Then frm.Schrittweite = 1
 frm.FDC.indDat = frm.FDC.indDat - frm.Schrittweite
 If frm.FDC.indDat < 0 Then frm.FDC.indDat = 0
' Call frm.FDC(frm.FDC.indDat).DateiLad(frm)
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doR�ckw�rtsCmd/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doR�ckw�rtsCmd

Function Key(KeyCode%, Shift%, frm As f�rIcon)
 Dim erg&
 On Error GoTo fehler
 If KeyCode = 27 Then
    If frm.K�rperTeil.ListCount > 0 Then
     erg = MsgBox("Wollen Sie wirklich abbrechen?", vbYesNo, "Sicherheitsr�ckfrage")
     If erg = vbNo Then Exit Function
    End If
    frm.Visible = False
    Call frm.ValidateControls
    On Error Resume Next
    frm.FDC.indDat = 0
    Unload frm
    End
 End If
' If KeyCode = 33 Then Call doR�ckw�rtsCmd(frm)
' If KeyCode = 34 Then Call doVorw�rtsCmd(frm) <- stellt den aktuellen Feldinhalt falsch ein!
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in key/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' key

' doErsterOffenerCmd
Function doersterOffenerCmd(frm As f�rIcon)
 On Error GoTo fehler
 frm.FDC.indDat = frm.FDC.ersterOffener()
 If frm.FDC.indDat > frm.FDC.Count Then frm.FDC.indDat = frm.FDC.Count
' Call frm.FDC(frm.FDC.indDat).DateiLad(frm)
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doersterOffenerCmd/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doersterOffenerCmd

Function doPatNameChange(frm As f�rIcon)
' Dim rAna As DAO.Recordset
 Dim rEi As New ADODB.Recordset
 Dim Inh$, zwi$
 Dim Auszug$, pos&
 Dim pos2&
' Static altPatId&
 On Error GoTo fehler
 Pat_id = 0
 If InStr(frm.PatName, "|") > 0 Then
  Pat_id = getPat_id(frm.PatName)
 End If
' If Pat_id <> altPatId Then
  frm.Fu�status = vNS
  frm.Doppler = vNS
  frm.Dopplerlabel = "Doppler"
  frm.Fu�statusBez = "Fu�status"
  If Pat_id <> 0 Then
   Set rEi = New ADODB.Recordset
   If QuelCStr = vNS Then
    Call ConstrFestleg(frm)
   End If
   zwi = QuelCStr
   Call rEi.Open("SELECT * FROM `eintraege` where pat_id = " & Pat_id & " and art like 'usdm%' ORDER BY zeitpunkt desc", zwi, adOpenKeyset, adLockReadOnly)
   Do
    If rEi.BOF Then Exit Do
    Inh = rEi!Inhalt
    pos = InStr(Inh, "A.tib.post.")
    pos2 = InStr(Inh, "aktuellen Blutdruck und ggf. Puls bitte extra eingeben")
    If pos > 0 Then
     Auszug = Mid(Inh, pos)
     If pos2 > 0 Then
      Auszug = Left(Auszug, pos2 - pos)
     End If
    End If
    frm.Fu�statusBez = "Pulsstatus vom " & rEi!zeitpunkt & ":"
    frm.Fu�status = REPLACE(REPLACE(Auszug, "A.tib.post.:", "A.t.p.:"), "Puls der re A.dors.ped.:", vbCrLf & "A.d.p.:")
    Exit Do
   Loop
   If Auszug = vNS Then
    rEi.Close
    zwi = QuelCStr
    Call rEi.Open("SELECT * FROM anamnesebogen where pat_id = " & Pat_id, zwi, adOpenKeyset, adLockReadOnly)
    If Not rEi.EOF Then
     If (Not IsNull(rEi("Puls Atp")) And rEi("Puls Atp") <> vNS) Or (Not IsNull(rEi("Puls Adp")) And rEi("Puls Atp") <> vNS) Then
      Auszug = "Atp:" & rEi("Puls Atp") & vbCrLf & rEi("Puls Adp")
      frm.Fu�statusBez = "Pulsstatus Anamnesebogen (vorgestellt am " & rEi!vorgestellt & "):"
      frm.Fu�status = Auszug
     End If
    End If
   End If
   rEi.Close
   zwi = QuelCStr
   frm.Doppler = vNS
   Call rEi.Open("SELECT * FROM `eintraege` where pat_id = " & Pat_id & " and art in (""doppler"",""duplex"") and inhalt not like ""%vene%"" and not inhalt like ""%halsschlag%"" and not inhalt like ""%caroti%"" ORDER BY zeitpunkt desc", zwi, adOpenKeyset, adLockReadOnly)
   If Not rEi.BOF Then
    frm.Dopplerlabel = UCase(Left(rEi!art, 1)) + Mid(rEi!art, 2) + " vom " + Format(rEi!zeitpunkt, "dd.mm.yy:")
    Do While Not rEi.EOF
     frm.Doppler = frm.Doppler + UCase(Left(rEi!art, 1)) + Mid(rEi!art, 2) + " " + Format(rEi!zeitpunkt, "dd.mm.yy:") + ": " + rEi!Inhalt + vbCrLf
     rEi.Move 1
    Loop
   End If
   rEi.Close
   zwi = QuelCStr
   frm.DTyp = vNS
   If Pat_id > 0 Then
    Call rEi.Open("SELECT dmtyp(" & Pat_id & ")", zwi, adOpenKeyset, adLockReadOnly)
    If Not rEi.BOF Then
     frm.DTyp = "Dm " & rEi.Fields(0)
    End If
   End If
'   altPatId = Pat_id
'  End If
 End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doPatNameChange/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doPatNameChange(f�ricon As F�rIcon)

Function doWieLetztesdPat(frm As f�rIcon, Optional obn�chster% = 0)
 Dim i&, von&, Bis&, schritt%, obgefunden%
 On Error GoTo fehler
 If frm.PatName <> vNS Then
  obgefunden = 0
  If obn�chster Then
    von = frm.FDC.indDat + 1
    Bis = frm.FDC.Count
    schritt = 1
  Else
    von = frm.FDC.indDat - 1
    Bis = 1
    schritt = -1
  End If
  For i = von To Bis Step schritt
   If frm.FDC(i).PatName = frm.PatName Then
    obgefunden = True
    frm.FDC(i).findeSatz
    Exit For
   End If
  Next i
  If Not obgefunden Then
   For i = frm.FDC.indDat + 1 To frm.FDC.Count
    If frm.FDC(i).PatName = frm.PatName Then
     obgefunden = True
     frm.FDC(i).findeSatz
     Exit For
    End If
   Next i
  End If
  If obgefunden Then
   Call doWieAnderes(frm)
  End If
 Else
  MsgBox "geht nicht, da Patientenname noch nicht ausgesucht"
 End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doWieLetztesdPat/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doWieLetztesdPat(frm As f�rIcon)

Function doWieN�chstes(frm As f�rIcon)
 On Error GoTo fehler
 If frm.FDC.indDat < frm.FDC.Count Then
  If frm.FDC(frm.FDC.indDat + 1).findeSatz Then
   Call doWieAnderes(frm)
  End If
 Else
  MsgBox "geht nicht, da Anfang"
 End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doWieVoriges/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doWieN�chstes

Function doWieVoriges(frm As f�rIcon)
 On Error GoTo fehler
 If frm.FDC.indDat > 1 Then
  If frm.FDC(frm.FDC.indDat - 1).findeSatz Then
   Call doWieAnderes(frm)
  End If
 Else
  MsgBox "geht nicht, da Ende"
 End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doWieVoriges/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doWieVoriges

Function doWieAnderes(frm As f�rIcon)
 Dim i%, zahl&, WT$
 On Error GoTo fehler
 If IsNull(rSteu!PatName) Then
  MsgBox "rsteu!patname = null, breche diese Aktion ab, vermutlich nach Daten�bertragung gleich gestartet!"
  Exit Function
 End If
  frm.PatName = rSteu!PatName
  frm.PatName.BackColor = &H80C0FF
  frm.Beschreibung = rSteu!Beschreibung
  frm.Beschreibung.BackColor = &H80C0FF
  frm.K�rperTeil = rSteu!Koerperteil
  frm.K�rperTeil.BackColor = &H80C0FF
  frm.obPat = -rSteu!verwendet
  If frm.obPat = 0 Then
   If Right(Trim(frm.PatName), 1) = ")" Then
    For i = Len(frm.PatName) To 1 Step -1
     If Mid(frm.PatName, i, 1) = " " Then
      On Error Resume Next
      zahl = Mid(frm.PatName, i + 2, Len(frm.PatName) - i - 2)
      On Error GoTo fehler
      If zahl > 0 Then
       frm.PatName = Left(frm.PatName, i - 1) + " (" + CStr(zahl + 1) + ")"
      Else
       frm.PatName = frm.PatName + " (1)"
      End If
      Exit For
     End If
    Next i
   Else
    frm.PatName = frm.PatName + " (1)"
   End If
  End If
  If Not IsNull(rSteu!WA) Then
  WT = WagnerText(frm, rSteu!WA)
   frm.Wagner = WT
   For i = 0 To frm.Wagner.ListCount - 1
    If frm.Wagner.List(i) = frm.Wagner.Text Then
'     frm.Wagner.ListIndex = i
     Exit For
    End If
   Next
'   frm.WagnerL = WT
   frm.Wagner.BackColor = &H80C0FF
'   frm.WagnerL.BackColor = &H80C0FF
   frm.Armstrong = ArmstrongText(frm, rSteu!WA)
   frm.Armstrong.BackColor = &H80C0FF
  End If
  Call frm.FDC(frm.FDC.indDat).findeSatz
  Exit Function
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doWieAnderes/" + App.Path)
  Case vbAbort: Call MsgBox("H�re auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' doWieAnderes

Public Function VerzSpei(frm As f�rIcon)
 Call fStSpei(HCU, RegStelle, "Quelle", frm.Quelle)
 Call fStSpei(HCU, RegStelle, "SteuerDB", frm.SteuerDB)
 Call fStSpei(HCU, RegStelle, "NamDB", frm.NamDB)
 Call fStSpei(HCU, RegStelle, "Kompr", frm.Kompr)
 Call fStSpei(HCU, RegStelle, "ArchPat", frm.ArchPat)
 Call fStSpei(HCU, RegStelle, "Archiv", frm.Archiv)
 Call fStSpei(HCU, RegStelle, "obAcc", frm.obAcc)
 Call fStSpei(HCU, RegStelle, "obMySQL", frm.obMySQL)
 Call fStSpei(HCU, RegStelle, "Kompressionsgrad", frm.Kompressionsgrad)
' Call fStSpei(HCU, RegStelle, "obMyQuelle1", frm.obMyQuelle1)
End Function ' VerzSpei(frm As f�rIcon)

Public Function do_Form_Unload(frm As f�rIcon)
 On Error GoTo fehler
 Call VerzSpei(frm)
 Call DForm_Unload(0)
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in do_Form_Unload/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Form_Unload

' kommt nur in do_Form_Unload vor
Public Function HolReg(frm As f�rIcon)
' Dim wsh As IWshShell_Class
 On Error GoTo fehler
 imAufbau = True
 If wsh Is Nothing Then Set wsh = New IWshShell_Class ' = CreateObject("Wscript.Shell")
 If FSO Is Nothing Then Set FSO = New FileSystemObject
 frm.Quelle = getReg(1, RegStelle, "Quelle")
 If Trim(frm.Quelle) = vNS Then frm.Quelle = sQuelle
 frm.SteuerDB = getReg(1, RegStelle, "SteuerDB")
 If Trim(frm.SteuerDB) = vNS Then frm.SteuerDB = sSteuerDB
 frm.NamDB = getReg(1, RegStelle, "NamDB")
 If Trim(frm.NamDB) = vNS Then frm.NamDB = sNamDB
 frm.Kompr = getReg(1, RegStelle, "kompr")
 If Trim(frm.Kompr) = vNS Then frm.Kompr = sKompr
 frm.ArchPat = getReg(1, RegStelle, "ArchPat")
 If Trim(frm.ArchPat) = vNS Then frm.ArchPat = sArchPat
 frm.Archiv = getReg(1, RegStelle, "Archiv")
 If Trim(frm.Archiv) = vNS Then frm.Archiv = sArchiv
 frm.Kompressionsgrad = getReg(1, RegStelle, "Kompressionsgrad")
 If LenB(Trim$(frm.Kompressionsgrad)) = 0 Then frm.Kompressionsgrad = sKompressionsgrad
 Dim zwi$
 zwi = getReg(1, RegStelle, "obAcc")
 If zwi = vNS Then
  frm.obAcc = False
  frm.obMySQL = True
 Else
  frm.obAcc = False ' zwi
  frm.obMySQL = True
 End If
' zwi = GetReg(1, RegStelle, "obMyQuelle")
' If zwi = vns Then
'  frm.obMyQuelle = True
' Else
'  frm.obMyQuelle = zwi
' End If
' zwi = GetReg(1, RegStelle, "obMyQuelle1")
' If zwi = vns Then
'  frm.obMyQuelle1 = False
' Else
'  frm.obMyQuelle1 = zwi
' End If
' If frm.obAcc = 0 And frm.obMyQuelle1 = 0 Then frm.obMyQuelle = 1
 imAufbau = False
' On Error Resume Next
' frm.AlteArchive = GetReg(1, RegStelle, "AlteArch")
' frm.obAlteArchive = wsh.RegRead("HKEY_CURRENT_USER" + "\" + RegStelle + "\" + "obAlteArch")
' frm.obQuelle = wsh.RegRead("HKEY_CURRENT_USER" + "\" + RegStelle + "\" + "obQuelle")
' frm.obEingel = wsh.RegRead("HKEY_CURRENT_USER" + "\" + RegStelle + "\" + "obEingel")
' frm.AlteArchive = GetReg(1, RegStelle, "AlteArch")
' frm.Eingel = GetReg(1, RegStelle, "Eingel")
Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in HolReg/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' HolReg

Public Function do_Form_Load(frm As f�rIcon)
 Dim Fld As Folder
 On Error GoTo fehler
 If FSO Is Nothing Then Set FSO = CreateObject("Scripting.Filesystemobject")
 frm.Status.Picture = LoadPicture(App.Path + "\..\icons\Mug of Tea.ico")
 IViewPfad = getIViewPfad()
 If IViewPfad = vNS Then
  MsgBox "IrfanView nach Registry offenbar nicht installiert. Dies wird aber ben�tigt. Breche ab."
  Unload frm
  End
 End If
 Call HolReg(frm)
 frm.Ausgabe = vNS
 Call f�lleLw(frm)
 'Dim QDB As DAO.Database
 
 'If QDB Is Nothing Then
 ' Set QDB = OpenDatabase(frm.NamDB)
  On Error Resume Next
  frm.Image1.Picture = LoadPicture("p:\Fotos Original\PICT0068.jpg") ', vbLPCustom, vbLPColor, 32, 32)
  frm.Picture2.Picture = LoadPicture("p:\Fotos Original\PICT0068.jpg") ', vbLPCustom, vbLPColor, 32, 32)
  On Error GoTo fehler
 'End If
' Set rNa = QDB.OpenRecordset("SELECT nachname + vns, vns + vorname + vns,*"" + format(gebdat,""D.M.YY"") + "" | ""+ cstr(pat_id) as T FROM namen ORDER BY nachname, vorname, gebdat;", dbOpenDynaset, ReadOnly)
' Do While Not rNa.EOF
'  frm!PatName.AddItem rNa!t
'  rNa.Move 1
' Loop
 
 Call doEnable(frm, 1)
 frm.WindowState = 2 ' ganz gro�
 Call Drehen.DForm_Load
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in do_Form_Load/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Form_Load

Public Function ConstrFestleg(frm As f�rIcon, Optional art%)
Const opti& = 2 + 4 '+ 8   ' 131118, 32 ' 1 + 2048 + 16384 + 131072
'ConStr$ = "DRIVER={MySQL ODBC 3.51 Driver};server=linux1;uid=praxis;pwd=sonne;option=" & opti
 Dim obmy%
 On Error GoTo fehler
 Select Case art
  Case 0
   If f�rIcon.obAcc Then obmy = False Else obmy = True
  Case 1
   obmy = False
  Case Is > 1
   obmy = True
 End Select
 Do
  If obmy = False Then
   QuelCStr = CStrAcc & sNamDB
   FotoCStr = CStrAcc & sSteuerDB
   obMySQL = 0
  Else
   Call frm.dbv.cnVorb(vNS, "anamnesebogen", "Quelle")
   QuelCStr = frm.dbv.CnStr
   Call frm.dbv.cnVorb(vNS, "jpg", "fotosinp")
   FotoCStr = frm.dbv.CnStr
   obMySQL = True
  End If
  On Error Resume Next
  Set FotoCn = Nothing
  FotoCn.Open FotoCStr
  If Err.Number = 0 Then
   On Error GoTo fehler
   Exit Do
  Else
   On Error GoTo fehler
  End If
 Loop
 frm.CnStr.Caption = frm.dbv.Constr
Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in do_Form_Load/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'ConstrFestleg(frm As f�rIcon)

' kommt vor in FDC_indnachWechsel und dowievoriges
Function ArmstrongText$(frm As f�rIcon, WA$)
 Dim i%
 If Len(WA) > 1 Then
  ArmstrongText = Mid(WA, 2, 1)
  For i = 0 To frm.Armstrong.ListCount - 1
   If UCase(ArmstrongText) = Left(frm.Armstrong.List(i), 1) Then
    ArmstrongText = frm.Armstrong.List(i)
    Exit For
   End If
  Next i
 End If
  Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ArmstrongText/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ArmstrongText

Function WagnerText$(frm As f�rIcon, WA$)
 Dim i%
 If Len(WA) > 0 Then
  WagnerText = Mid(WA, 1, 1)
  For i = 0 To frm.Wagner.ListCount - 1
   If WagnerText = Left(frm.Wagner.List(i), 1) Then
    WagnerText = frm.Wagner.List(i)
    Exit For
   End If
  Next i
 End If
  Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in WagnerText/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' WagnerText

Public Sub do_Start(frm As f�rIcon)
 On Error GoTo fehler
'Function do_main(Optional frm As f�rIcon)
 Dim cont$, T1, T2, ZNr%, j&, Zus%, Dt, daz%, neuz%, erge, uBo&
 Dim FName$, NeuName$, ArchPatName$, Fil As File, Wav As File, WavDat$
 On Error GoTo fehler
 If Right(frm.Quelle, 1) <> "\" Then frm.Quelle = frm.Quelle + "\"
 If Right(frm.ArchPat, 1) <> "\" Then frm.ArchPat = frm.ArchPat + "\"
 Call VerzSpei(frm)
 Call frm.DBInit
' If frm Is Nothing Then Set frm = f�rIcon
 frm.Status.Picture = LoadPicture(App.Path + "\..\icons\info.ico")
 
 If frm.Lw.ListIndex < frm.Lw.ListCount - 1 Then
  Call DatKop(frm, Left(frm.Lw.List(frm.Lw.ListIndex), 1))
 End If
 
 Set frm.FDC = New FDateiColl
 Call frm.FDC.Init
 frm.FDC.indDat = frm.FDC.ersterOffener
 If frm.FDC.indDat <> 0 Then
  Call doEnable(frm, Status:=2)
  frm.DSZahl = frm.FDC.Count
 'Call frm.FDC(frm.FDC.indDat).aktSpeichern(frm)
  
  frm.Status.Picture = LoadPicture(App.Path + "\..\icons\Mug of Tea.ico")
  DoEvents
  Call ConstrFestleg(frm)
  Call Auswahlen(frm)
 Else
  Call doAusgabe("Keine Fotos")
  frm.Status.Picture = LoadPicture(App.Path + "\..\icons\Mug of Tea.ico")
 End If
 Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in do_Start/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' do_Start

Sub Auswahlen(frm As f�rIcon)
 Dim rNaA As New ADODB.Recordset
 On Error GoTo fehler
 Call AuswName(frm)
 frm!K�rperTeil.AddItem "Fu� re"
 frm!K�rperTeil.AddItem "Fu� li"
 frm!K�rperTeil.AddItem "Zehen re"
 frm!K�rperTeil.AddItem "Zehen li"
 frm!K�rperTeil.AddItem "Ferse re"
 frm!K�rperTeil.AddItem "Ferse li"
 frm!K�rperTeil.AddItem "US re"
 frm!K�rperTeil.AddItem "US li"
' If SDB Is Nothing Then Set SDB = OpenDatabase(frm.SteuerDB)
' rNaA.Close
 rNaA.Open "SELECT distinct k�rperteil as k FROM `jpg` where not isnull(k�rperteil) ORDER BY k�rperteil", FotoCn, adOpenStatic, adLockReadOnly
 Do While Not rNaA.EOF
  frm!K�rperTeil.AddItem rNaA!k
  rNaA.Move 1
 Loop
 
 rNaA.Close
 rNaA.Open "SELECT distinct beschreibung as b FROM `jpg` where not isnull(beschreibung) ORDER BY beschreibung", FotoCn, adOpenStatic, adLockReadOnly
 Do While Not rNaA.EOF
  frm!Beschreibung.AddItem rNaA!b
  rNaA.Move 1
 Loop
 
 Dim i%, j%, ctl As ComboBox
 For i = 1 To 2
  Select Case i
   Case 1: Set ctl = frm!Wagner
'   Case 2: Set Ctl = frm!WagnerL
  End Select
  With ctl
  For j = .ListCount - 1 To 0 Step -1
   .RemoveItem (j)
  Next j
 .AddItem "- kein Wagnerstadium"
 .AddItem "0 pr�- oder postulcerative L�sion"
 .AddItem "1 oberfl�chliche Wunde"
 .AddItem "2 Wunde bis Ebene Sehne/Kapsel"
 .AddItem "3 Wunde bis Ebene Knochen/Gelenk"
 .AddItem "4 Nekrose Fu�teile"
 .AddItem "5 Nekrose Fu� ganz"
  End With
 Next
 
 frm!Armstrong.AddItem "- nicht anzuwenden"
 frm!Armstrong.AddItem "A ohne Infektion/Isch�mie"
 frm!Armstrong.AddItem "B mit Infektion"
 frm!Armstrong.AddItem "C mit Isch�mie"
 frm!Armstrong.AddItem "D mit Infektion+Isch�mie"
 
 Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Auswahlen/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Auswahlen
 
Function AuswName(frm As f�rIcon)
 Dim zwi$, sql$
 On Error GoTo fehler
 'Dim cnADO As New ADODB.Connection
 'cnADO.Open ConStr
 Dim rNaA As New ADODB.Recordset
 
 zwi = QuelCStr
 If InStr(zwi, "MySQL") > 0 Or InStr(zwi, "MSDASQL") > 0 Then
  sql = "SELECT distinct n.nachname, n.vorname, n.gebdat, n.pat_id FROM namen n LEFT JOIN faelle f USING (pat_id) WHERE bhfb > SUBDATE(now(),INTERVAL IF(schgr=90,24,9) MONTH) ORDER BY nachname, vorname, gebdat"
 Else
  sql = "SELECT distinct n.nachname, n.vorname, n.gebdat, n.pat_id FROM namen n LEFT JOIN faelle f ON n.pat_id = f.pat_id WHERE bhfb > now()- IIF(schgr=90,730,270) ORDER BY n.nachname, n.vorname, n.gebdat"
 End If
' sql = "SELECT distinct nachname, vorname, gebdat, pat_id FROM namen ORDER BY nachname, vorname, gebdat"
 rNaA.Open sql, zwi, adOpenStatic, adLockReadOnly
 zwi = frm.PatName
 frm.PatName.Clear
 frm.PatName = zwi
 Dim t$
 Do While Not rNaA.EOF
  t = rNaA!nachname + ", " + rNaA!vorname + ",*" + Format(rNaA!gebdat, "D.M.YY") + " | " + CStr(rNaA!Pat_id)
  frm!PatName.AddItem t
  rNaA.Move 1
 Loop
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in AuswName/" + App.Path)
 Case vbAbort: Call MsgBox("H�re auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' AuswName
