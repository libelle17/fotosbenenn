VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FDatei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Datei"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvFil As File 'lokale Kopie
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
'Private mvBM As String 'Bookmark
Private altPfad$ ' Funktion von bookmark unter Ado -> eindeutiger Index Pfad
Private mvGelöscht As Boolean 'lokale Kopie
Private mvBearbeitet As Boolean 'lokale Kopie
Private mvVerwendet As Boolean 'lokale Kopie
'Private mvobAngezeigt As Boolean 'lokale Kopie
Private mvNeuerName As String 'lokale Kopie
Private mvNNgelöscht As Boolean 'lokale Kopie
Private mvWavPfad As String 'lokale Kopie
Private mvWavErstellt As Date 'lokale Kopie
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvWavGröße As Long 'lokale Kopie
Private mvWavGelöscht As Boolean 'lokale Kopie
Private mvPict As Long 'lokale Kopie
Private mvKörperteil As String 'lokale Kopie
Private mvBeschreibung As Variant 'lokale Kopie
Private mvPatDatum As Date 'lokale Kopie
Private mvWA As String 'lokale Kopie
Private mvPat_ID As Long 'lokale Kopie
Private mvPatName As String 'lokale Kopie
Private mvHelligkeit As Long
Private mvKontrast As Long
Private mvGamma As Long
Public Event Ausgabe(Str$)
Private mvarIndex As Long 'lokale Kopie
Dim FSO As New FileSystemObject


Public Property Let Index(ByVal vData As Long)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.index = 5
    mvarIndex = vData
End Property

Public Property Get Index() As Long
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.index
    Index = mvarIndex
End Property

Public Function istKleinerAls(ByRef E2) As Boolean
 If Fil.name < E2.name Then
  istKleinerAls = -1
 Else
  istKleinerAls = 0
 End If
End Function

Public Function unGleich(ByRef E2) As Boolean
 If Fil.name <> E2.name Then unGleich = True
End Function

Private Function doAusgabe(Str$)
 RaiseEvent Ausgabe(Str)
End Function ' doAusgabe

'Public obupd%
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
'Um dieses Ereignis auszulösen, verwenden Sie RaiseEvent mit der folgenden Syntax:
'RaiseEvent zublättern[(arg1, arg2, ... , argn)]
'Public Event zublättern()
'Um dieses Ereignis auszulösen, verwenden Sie RaiseEvent mit der folgenden Syntax:
'RaiseEvent aufblättern[(arg1, arg2, ... , argn)]
'Public Event aufblättern()
Function machNeuenNamen$()
' frm kann fürIcon oder rSteu sein
  On Error GoTo fehler
  machNeuenNamen = vNS
  If Trim(mvPatName) <> vNS Then machNeuenNamen = Trim(Split(mvPatName, "*")(0))
  If Right(machNeuenNamen, 1) = "," Then machNeuenNamen = Left(machNeuenNamen, Len(machNeuenNamen) - 1)
  On Error Resume Next
  On Error GoTo fehler
  If mvVerwendet Then
   machNeuenNamen = machNeuenNamen & ", Foto "
  Else
   If mvKörperteil <> vNS Then machNeuenNamen = machNeuenNamen & "; "
  End If
  machNeuenNamen = machNeuenNamen & mvKörperteil
  If Trim(machNeuenNamen) <> vNS Then
      ' 10.9.19 DateCreate durch DateLastModified ersetzt
   machNeuenNamen = machNeuenNamen & Format$(Fil.DateLastModified, " dd.mm.yy ")
   machNeuenNamen = machNeuenNamen & mvBeschreibung
   If Trim(mvWA) <> vNS Then machNeuenNamen = machNeuenNamen & " WA " & mvWA
  End If
  machNeuenNamen = REPLACE$(REPLACE$(REPLACE$(machNeuenNamen, ":", " - "), "\", " - "), "/", " - ")
  Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in machNeuenNamen/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' machNeuenNamen(frm As FürIcon)

' Vergleicht ein Datenbankfeld mit der zugehörigen Klasseneigenschaft
Private Sub vglS(DBEigs, ClassEigs, rspfad$, obupd%, Optional obPfad%)
' Dim afn%
 Dim sql$
 'rSteu.Open "SELECT * FROM `jpg` where pfad = '" & Replace(altPfad, "\", "\\") & "'", FotoCn, adOpenDynamic, adLockOptimistic
 If IsNull(ClassEigs) Then
  MsgBox "Fehler in vgls bei Vergleich der Datenbankeigenschaft '" & DBEigs & "' mit NULL bei '" & rspfad & "'"
  Exit Sub
 End If
 On Error GoTo fehler
 Dim para$
 If VarType(ClassEigs) = vbDate Then
  para = datForm(ClassEigs)
 ElseIf VarType(ClassEigs) = vbBoolean Then
  para = "'" & --ClassEigs & "'"
 Else
  If obPfad Then
   para = "'" & REPLACE(REPLACE(REPLACE(ClassEigs, "'", "''"), "<", " kl. "), ">", " gr. ") & "'"
   If InStr(FotoCn, "MySQL") > 0 Or InStr(FotoCn, "MSDASQL") > 0 Then
    para = REPLACE(para, "\", "\\")
   End If
  Else
   para = "'" & REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(ClassEigs, "'", "''"), "<", " kl. "), ">", " gr. "), ":", vNS), "\", "/") & "'"
  End If
 End If
 Dim afnl&
 'Call FotoCn.Execute("update `jpg` set `" & DBEigs & "` = " & para & " where pfad = '" & IIf(InStr(FotoCn, "MySQL") > 0 Or InStr(FotoCn, "MSDASQL") > 0, REPLACE(rspfad, "\", "\\"), rspfad) & "' and (isnull(`" & DBEigs & "`) or not `" & DBEigs & "` = " & para & ")", afn)
 Dim fex As ADODB.Recordset
 myFrag fex, "UPDATE `jpg` SET `" & DBEigs & "` = " & para & " WHERE pfad = '" & IIf(InStr(FotoCn, "MySQL") > 0 Or InStr(FotoCn, "MSDASQL") > 0, REPLACE(rspfad, "\", "\\"), rspfad) & "' AND (ISNULL(`" & DBEigs & "`) OR NOT `" & DBEigs & "` = " & para & ")", , FotoCn, , , afnl
 If afnl <> 0 Then
  obupd = -1
 End If
 Exit Sub
' On Error Resume Next
' If (IsNull(rSteu.Fields(DBEigs)) And Not IsNull(ClassEigs)) Or (rSteu.Fields(DBEigs) <> ClassEigs) Then
' On Error GoTo fehler
'  obupd = -1
'  rSteu.Fields(DBEigs) = ClassEigs
'  Dim nVar
'  SELECT Case VarType(ClassEigs)
'   Case vbBoolean: nVar = ClassEigs + 0
'   Case vbString: nVar = "'" & ClassEigs & "'"
'   Case vbDate: nVar = datForm(ClassEigs)
'   Case Else: nVar = ClassEigs
'  End Select
'  sql = "update `jpg` set `" & DBEigs & "` = " & nVar & " where pfad = '" & Replace(rspfad, "\", "\\") & "'"
'  Call FotoCn.Execute(sql, afn)
'  If afn <> 1 Then
'   Stop
'  End If
' End If
' Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in vglS/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub      ' vglS

' ermittelt die vierstellige Nummer nach "Pict" im Dateinamen
Private Function getPict&(Datei$)
 Dim Pos1&
 On Error GoTo fehler
 Call DateiArt(Datei, getPict)
' Pos1 = InStr(UCase(Datei), "PICT")
' If Pos1 = 0 Then Pos1 = InStr(UCase(Datei), "BILD")
' If Pos1 = 0 Then Pos1 = InStr(UCase(Datei), "IMG_")
' If Pos1 > 0 Then
'  getPict = CLng(Mid(Datei, Pos1 + 4, 4))
' End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in WNetKorr/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' getPict

Function findeSatz%()
 Dim DPfad$, runde&
 DPfad = mvFil.Path
 On Error GoTo fehler
 If FotoCStr = vNS Then
  Call ConstrFestleg(fürIcon)
 End If
 Set rSteu = Nothing
 rSteu.Open "SELECT Pfad, NNgelöscht NNgeloescht, bearbeitet, verwendet, KörperTeil Koerperteil, Beschreibung, WA, PatName, Helligkeit, Kontrast, Gamma, Pict, PatDatum, NeuerName FROM fotosinp.jpg WHERE pfad = '" & IIf(obMySQL, REPLACE(DPfad, "\", "\\"), DPfad) & "' AND NOT gelöscht AND NOT NNgelöscht", FotoCn, adOpenStatic, adLockOptimistic
 If Not rSteu.BOF Then
  findeSatz = True
 End If
 Exit Function
fehler:
If Err.Number = -2147217887 Then ' gleichzeitiger Import
 runde = runde + 1
 If runde < 3 Then
  FotoCn.Close
  FotoCn.Open
  Resume
 End If
End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in findeSatz/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): FotoCn.Close: FotoCn.Open: Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' findeSatz%

' Findet den zu einer Datei gehörigen Dateinamen
'Function findeSatzAlt%(rspfad$)
'  Dim DName$
'  DName = mvFil.Name
''  If DName Like "*######## ######.*" Then DName = Left(DName, Len(DName) - 20) & Right(DName, 4) ' 25.5.08
'  On Error GoTo fehler
''  Call doDBInit
''  rSteu.Seek "=", DName
'  'rSteu.Find "Name = " & DName, 0, adSearchForward, 1
'  If FotoCStr = vns Then
'   Call ConstrFestleg(fürIcon)
'  End If
'  Set rSteu = Nothing
'  rSteu.Open "SELECT * FROM `jpg` where name = '" & DName & "'", FotoCn, adOpenDynamic, adLockOptimistic
'    Do
'     If rSteu.BOF Then Exit Do
'     If rSteu.EOF Then Exit Do
''     If rSteu.NoMatch Then Exit Do
''     If LCase(rSteu!Name) <> LCase(DName) Then Exit Do
'     If Not rSteu!gelöscht And Not rSteu!NNgelöscht Then
'      findeSatzAlt = -1
'      rspfad = rSteu!Pfad
'      Exit Do
'     End If
'     rSteu.Move 1
'    Loop
'    Exit Function
'fehler:
'SELECT Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vns, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in findeSatzalt/" + App.Path)
' Case vbAbort: Call MsgBox("Höre auf"): End
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End Select
'End Function ' findeSatz
Private Function fnN$(Var)
 If IsNull(Var) Then
  fnN = vNS
 Else
  fnN = Var
 End If
End Function ' fnN

Private Function fnS&(Var, Def&)
 If IsNull(Var) Then
  fnS = Def
 Else
  fnS = Var
 End If
End Function ' fnN

Public Sub LadDatenSatz() ' von der Datenbank in die Variable FDatenColl
    Dim obupd%, Wav As File, rAf&
    Dim rspfad$, sql$
    Dim i&
    Dim rs As New ADODB.Recordset
    On Error GoTo fehler
wieder:
    If findeSatz Then
     rspfad = rSteu!Pfad
     altPfad = rspfad ' eindeutiger Index
     obupd = 0
     mvNNgelöscht = rSteu!NNgeloescht ' dürfte nicht vorkommen
     mvBearbeitet = rSteu!bearbeitet
     mvVerwendet = rSteu!verwendet
     mvKörperteil = fnN(rSteu!Koerperteil)
     mvBeschreibung = fnN(rSteu!Beschreibung)
     mvWA = fnN(rSteu!WA)
     mvPatName = fnN(rSteu!PatName)
     mvHelligkeit = fnS(rSteu!Helligkeit, 90)
     mvKontrast = fnS(rSteu!Kontrast, 90)
     mvGamma = fnS(rSteu!Gamma, 250)
     mvPict = fnS(rSteu!Pict, 0)
    Else
     rspfad = mvFil.Path
     Call FotoCn.Execute("insert into `jpg`(name,verwendet,pfad) values('" & IIf(InStr(FotoCn, "MySQL") > 0 Or InStr(FotoCn, "MSDASQL") > 0, REPLACE(mvFil.name, "\", "\\"), mvFil.name) & "'," & IIf(InStr(FotoCn, "MySQL") > 0 Or InStr(FotoCn, "MSDASQL") > 0, 1, -1) & ",'" & IIf(InStr(FotoCn, "MySQL") > 0 Or InStr(FotoCn, "MSDASQL") > 0, REPLACE(rspfad, "\", "\\"), rspfad) & "')", rAf)
     If rAf = 1 Then
      Set rs = Nothing
      rs.Open "SELECT last_insert_id() id", FotoCn, adOpenStatic, adLockReadOnly
      Dim info&, tmpfil$, txt$, spli$()
      tmpfil = "c:\tmpimginfo" & Format(Now(), "yyyyMMddhhmmss") & ".txt"
      info = Shell(vNS & IViewPfad & " """ & rspfad & """ /fullinfo /info=" & tmpfil)
      Call WartAufProzeß(info)
      If FSO.FileExists(tmpfil) Then
       Open tmpfil For Input As #377
       Do While Not EOF(377)
        Input #377, txt
        SplitNeuArr txt, Array(" = "), spli
        If UBound(spli) <> 1 Then SplitNeuArr txt, Array(" - "), spli
        If UBound(spli) = 1 Then
         On Error Resume Next
         For i = 1 To 3
          Err.Clear
          FotoCn.Execute "update `jpg` set `" & spli(0) & "` = '" & fUmwfSQL(spli(1)) & "' where id = " & rs!ID, rAf
          If Err.Number = 0 Then Exit For
          Err.Clear
          If i = 1 Then
           FotoCn.Execute "alter table `jpg` modify `" & spli(0) & "` varchar(" & Len(fUmwfSQL(spli(1))) & ")", rAf
          ElseIf i = 2 Then
           FotoCn.Execute "alter table `jpg` add `" & spli(0) & "` varchar(" & Len(fUmwfSQL(spli(1))) & ")", rAf
          End If
         Next i
        End If
       Loop
       Close #377
       Kill tmpfil
      End If
     End If
'     rSteu.AddNew
'     rSteu!Name = mvFil.Name
'     rSteu!verwendet = -1
'     rspfad = mvFil.Path
'     rSteu!Pfad = rspfad
'     rSteu.Update
     obupd = -1
    End If
    Call vglS("erstellt", mvFil.DateCreated, rspfad, obupd)
    Call vglS("geändert", mvFil.DateLastModified, rspfad, obupd)
    Call vglS("Größe", mvFil.Size, rspfad, obupd)
'    Call vglS("Pfad", mvFil.Path, rsPfad, obupd)
    Dim obPatDatum%
    obPatDatum = 0
    If rSteu.BOF Then
     obPatDatum = True
    Else
     If IsNull(rSteu!PatDatum) Then
      obPatDatum = -1
     Else
      If rSteu!PatDatum = 0 Then
       obPatDatum = -1
      End If
     End If
    End If
    If obPatDatum Then
    ' 10.9.19 DateCreate durch DateLastModified ersetzt
     Call vglS("PatDatum", mvFil.DateLastModified, rspfad, obupd)       'DateValue(rSteu!erstellt)
    End If
    Dim obPict%
    obPict = 0
    If rSteu.BOF Then
     obPict = True
    Else
     If IsNull(rSteu!Pict) Then
      obPict = -1
     Else
      If rSteu!Pict = 0 Then
       obPict = -1
      End If
     End If
    End If
    If obPict Then
     mvPict = getPict(mvFil.name)
     Call vglS("Pict", mvPict, rspfad, obupd)
    End If
'    mvWavPfad = Replace(LCase(mvFil.Path), ".jpg", ".wav")
    Set Wav = GetWav
    If Not Wav Is Nothing Then
'     mvWavPfad = Left$(mvFil.Path, Len(mvFil.Path) - Len(mvFil.Name)) & mvWavPfad
'     Set Wav = FSO.GetFile(mvWavPfad)
     mvWavPfad = Wav.Path
     Call vglS("WavPfad", mvWavPfad, rspfad, obupd, obPfad:=True)
     mvWavErstellt = Wav.DateCreated
     Call vglS("WavErstellt", mvWavErstellt, rspfad, obupd)
     mvWavGröße = Wav.Size
     Call vglS("WavGröße", mvWavGröße, rspfad, obupd)
     mvWavGelöscht = 0
    Else
     mvWavGelöscht = -1
    End If
    Call vglS("WavGelöscht", mvWavGelöscht, rspfad, obupd)
    Call vglS("gelöscht", False, rspfad, obupd)
'    If IsNull(rSteu!NeuerName) Or rSteu!NeuerName = vns And rSteu!Pict <> 0 Then
'     Dim rs As New ADODB.Recordset
''     Set rs = SDB.OpenRecordset("SELECT * FROM `jpg` where pict = " + CStr(rSteu!Pict) + " and not isnull(neuername);")
'     sql = "SELECT * FROM `jpg` where pict = " + CStr(rSteu!Pict) + " and not isnull(neuername);"
'     rs.Open sql, FotoCn, adOpenDynamic, adLockOptimistic
'     If Not rs.BOF Then
'      rspfad = rs!Pfad
'      Call vglS("NeuerName", rs!NeuerName, rspfad, obupd)
'      Call vglS("KörperTeil", rs!Koerperteil, rspfad, obupd)
'      Call vglS("Beschreibung", rs!Beschreibung, rspfad, obupd)
'      Call vglS("verwendet", rs!verwendet, rspfad, obupd)
'      If Not IsNull(rs!WA) Then Call vglS("WA", rs!WA, rspfad, obupd)
'      Call vglS("PatDatum", rs!PatDatum, rspfad, obupd)
'      If Not IsNull(rs!Pat_id) Then Call vglS("Pat_id", rs!Pat_id, rspfad, obupd)
'     End If
'    End If
    If Not rSteu.EOF Then If Not IsNull(rSteu!NeuerName) Then mvNeuerName = rSteu!NeuerName
  If obupd Then
'   rSteu.CancelUpdate
'   GoTo wieder
'   rSteu.Update
   
'   mvBM = rSteu.LastModified
   Set rSteu = Nothing
   rSteu.Open "SELECT Pfad, NNgelöscht NNgeloescht, bearbeitet, verwendet, KörperTeil Koerperteil, Beschreibung, WA, PatName, Helligkeit, Kontrast, Gamma, Pict, PatDatum, NeuerName FROM `jpg` where pfad = '" & IIf(InStr(FotoCn, "MySQL") > 0 Or InStr(FotoCn, "MSDASQL") > 0, REPLACE(altPfad, "\", "\\"), altPfad) & "'", FotoCn, adOpenDynamic, adLockOptimistic
  Else
'   rSteu.CancelUpdate
  End If
  Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in LadDatenSatz/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub  ' LadDatenSatz

Public Function GetWav() As File
 Dim Pfad$
 Pfad = Dir(Left$(mvFil.Path, Len(mvFil.Path) - 4) & "*.wav")
 If LenB(Pfad) = 0 Then
  Pfad = Dir(Left$(mvFil.Path, Len(mvFil.Path) - Len(mvFil.name)) & Left$(mvFil.name, 8) & "*.wav")
 End If
 If LenB(Pfad) = 0 Then
  Pfad = Dir(Left$(mvFil.Path, Len(mvFil.Path) - Len(mvFil.name)) & "SND" & Mid$(mvFil.name, 4, 5) & "*.wav")
 End If
 If LenB(Pfad) <> 0 Then
  Set GetWav = FSO.GetFile(Left$(mvFil.Path, Len(mvFil.Path) - Len(mvFil.name)) & Pfad)
 End If
End Function ' GetWav

Public Sub SchreibDatenSatz()
  Dim obupd%, rspfad$
  Dim cons$
  Dim testAdo As ADODB.Recordset
  Dim tAdo As New ADODB.Recordset
  Dim i%, pka%
  rspfad = altPfad
  On Error Resume Next
'  rSteu.Bookmark = mvBM
  Set rSteu = Nothing
  Err.Clear
  rSteu.Open "SELECT Pfad, NNgelöscht NNgeloescht, bearbeitet, verwendet, KörperTeil Koerperteil, Beschreibung, WA, PatName, Helligkeit, Kontrast, Gamma, Pict, PatDatum, NeuerName FROM `jpg` WHERE pfad = '" & IIf(InStr(FotoCn, "MySQL") > 0 Or InStr(FotoCn, "MSDASQL") > 0, REPLACE(altPfad, "\", "\\"), altPfad) & "'", FotoCn, adOpenDynamic, adLockOptimistic
  If Err.Number <> 0 Then
   On Error GoTo fehler
   cons = FotoCn.ConnectionString
   FotoCn.Close
   FotoCn.Open FotoCn.ConnectionString & "pwd=" & p1 & "nne" & ";"
   rSteu.Open "SELECT Pfad, NNgelöscht NNgeloescht, bearbeitet, verwendet, KörperTeil Koerperteil, Beschreibung, WA, PatName, Helligkeit, Kontrast, Gamma, Pict, PatDatum, NeuerName FROM `jpg` WHERE pfad = '" & IIf(InStr(FotoCn, "MySQL") > 0 Or InStr(FotoCn, "MSDASQL") > 0, REPLACE(altPfad, "\", "\\"), altPfad) & "'", FotoCn, adOpenDynamic, adLockOptimistic
  End If
  On Error GoTo fehler
'  rSteu.Edit
  If Not rSteu.BOF Then
   Call vglS("gelöscht", mvGelöscht, rspfad, obupd)
   Call vglS("bearbeitet", -1, rspfad, obupd)
   Call vglS("verwendet", mvVerwendet, rspfad, obupd)
'  Call vglS(rSteu!NNgelöscht, mvNNgelöscht, rsPfad, obupd)
   If mvPatName <> vNS Then
    Do
     mvNeuerName = machNeuenNamen
     Set tAdo = Nothing
     tAdo.Open "SELECT * FROM `jpg` WHERE NeuerName = '" & REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(mvNeuerName, "'", "''"), "<", " kl. "), ">", " gr. "), ":", vNS), "\", "/") & "' and pfad <> '" & IIf(InStr(FotoCn, "MySQL") > 0 Or InStr(FotoCn, "MSDASQL") > 0, REPLACE(altPfad, "\", "\\"), altPfad) & "'", FotoCn, adOpenStatic, adLockReadOnly
     If tAdo.BOF Then Exit Do
     pka = 1
     If Trim(mvKörperteil) Like "*(*)" Then
      For i = Len(Trim(mvKörperteil)) To 1 Step -1
       If Mid(Trim(mvKörperteil), i, 1) = "(" Then
        pka = Mid(Trim(mvKörperteil), i + 1, Len(Trim(mvKörperteil)) - i - 1)
        pka = pka + 1
        Exit For
       End If
      Next i
     End If
     mvKörperteil = Trim(Left(mvKörperteil, IIf(i = 0, Len(mvKörperteil), i - 1))) & " (" & pka & ")"
    Loop
   End If
   Call vglS("NeuerName", mvNeuerName, rspfad, obupd)
'  Call vglS(rSteu!WavPfad, mvWavPfad, rsPfad, obupd)
   Call vglS("KörperTeil", mvKörperteil, rspfad, obupd)
   Call vglS("Beschreibung", mvBeschreibung, rspfad, obupd)
   Call vglS("WA", mvWA, rspfad, obupd)
   Call vglS("PatName", mvPatName, rspfad, obupd)
   Call vglS("PatDatum", mvPatDatum, rspfad, obupd)
   Call vglS("Helligkeit", mvHelligkeit, rspfad, obupd)
   Call vglS("Kontrast", mvKontrast, rspfad, obupd)
   Call vglS("Gamma", mvGamma, rspfad, obupd)
   If InStr(PatName, "|") > 0 Then
    mvPat_ID = getPat_id(mvPatName)
    Call vglS("Pat_id", mvPat_ID, rspfad, obupd)
   End If
  End If
  If obupd Then
'   rSteu.Update
'   mvBM = rSteu.LastModified
   altPfad = rSteu!Pfad
  Else
'   rSteu.CancelUpdate
  End If
'  rSteu.Bookmark = mvBM
  Set rSteu = Nothing
  rSteu.Open "SELECT Pfad, NNgelöscht NNgeloescht, bearbeitet, verwendet, KörperTeil Koerperteil, Beschreibung, WA, PatName, Helligkeit, Kontrast, Gamma, Pict, PatDatum, NeuerName FROM `jpg` where pfad = '" & IIf(InStr(FotoCn, "MySQL") > 0 Or InStr(FotoCn, "MSDASQL") > 0, REPLACE(altPfad, "\", "\\"), altPfad) & "'", FotoCn, adOpenDynamic, adLockOptimistic
  Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in LadDatenSatz/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' SchreibDatenSatz ' rSteuKoord

Function doÜbertragen(Kompr$, ArchPat$, Archiv$, Hel&, Kontr&, Gam&, KG$, Optional stehenLassen%)
   Dim erg&, NeuerName$, Ziel$, Var$, obVar%
    Dim mDN$, mGesN$, lfdnr&, direrg&
    Dim DErg$, DArt%
   On Error GoTo fehler
   If findeSatz Then
'    Stop
    If IsNull(rSteu!PatName) Then Exit Function
    If Trim(rSteu!PatName) = vNS Then Exit Function
    If rSteu!verwendet Then
     If IsNull(rSteu!Koerperteil) Then Exit Function
     If rSteu!Koerperteil = vNS Then Exit Function
     If IsNull(rSteu!Beschreibung) Then Exit Function
     If rSteu!Beschreibung = vNS Then Exit Function
    End If
    NeuerName = FSO.GetFolder(Kompr)
    If Right(NeuerName, 1) <> "\" Then NeuerName = NeuerName + "\"
    mDN = machNeuenNamen()
    mGesN = NeuerName & mDN & ".jpg"
    lfdnr = 0
    Do
     DErg = Dir(mGesN)
     If DErg = vNS Then Exit Do
     lfdnr = lfdnr + 1
     mGesN = NeuerName & mDN & "_" & lfdnr & ".jpg"
    Loop
    Dim DateiPfad$
    DateiPfad = mvFil.Path
    If mvHelligkeit <> 90 Or mvKontrast <> 90 Or mvGamma <> 250 Then
     Var = getVariantePfad(mvFil)
     obVar = FSO.FileExists(Var)
     If obVar Then DateiPfad = Var Else DateiPfad = mvFil.Path
    End If
    If FSO.FileExists(DateiPfad) Then
     erg = Komprimiere(DateiPfad, mGesN, Hel, Kontr, Gam, KG, obVar)
    Else
     erg = 0
    End If
    If FSO.FileExists(mGesN) And stehenLassen = 0 Then
     If rSteu!verwendet Then
      Ziel = FSO.GetFolder(ArchPat)
     Else
      On Error Resume Next
      Err.Clear
      Ziel = FSO.GetFolder(Archiv)
      If Err.Number <> 0 Then
       Ziel = FSO.GetFolder("P:\Fotos fachfremd")
      End If
      On Error GoTo fehler
     End If
     If Right(Ziel, 1) <> "\" Then Ziel = Ziel + "\"
     On Error Resume Next
     Dim mvName$, mvPfad$, wavPfad$, WavFil As File
     mvName = mvFil.name
     mvPfad = mvFil.Path
     Set WavFil = Me.GetWav
     DArt = DateiArt(WavFil.name)
     wavPfad = VerschiebeFD(WavFil.Path, Ziel & WavFil.name, DArt, True)
     Call vglS("wavpfad", wavPfad, mvPfad, True, True)
     On Error GoTo fehler
     DArt = DateiArt(mvName)
     Call VerschiebeFD(mvPfad, Ziel & mvName, DArt)
     If FSO.FileExists(Ziel & mvName) Then
      Call vglS("tpfad", Ziel & mvName, mvPfad, True, True)
     End If
     If obVar Then
      DArt = DateiArt(Var)
      Call VerschiebeFD(Var, Ziel & Dir(Var), DArt)
     End If
    End If
   End If
   DoEvents
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doÜbertragen/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doÜbertragen

' kommt vor in Rotier, cmdSaveAsRotated_jpg
Function LöscheDatei(D1$, Optional unsicher%)
 On Error GoTo fehler
 If unsicher Then
  Call doAusgabe("Versuche zu Löschen: ")
  On Error Resume Next
 Else
  Call doAusgabe("Lösche: ")
 End If
 Call doAusgabe(D1)
 Call FSO.DeleteFile(D1)
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in LöscheDatei/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LöscheDatei

Function VerschiebeFD$(D1$, D2$, Optional DArt = -1, Optional unsicher%)
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
 If DArt <> 0 Then ' UCase(D2a) Like "*BILD####*" Or UCase(D2a) Like "*PICT####*" Or UCase(D2a) Like "*IMG_####*" Then
  If Not D2a Like "*######## ######*" Then
   If D2a Like "*.???" Then
    D2a = Left(D2a, Len(D2a) - 4) & Format(FileDateTime(D2a), " yyyymmdd hhmmss") & Right(D2a, 4)
   ElseIf D2a Like "*.??" Then
    D2a = Left(D2a, Len(D2a) - 3) & Format(FileDateTime(D2a), " yyyymmdd hhmmss") & Right(D2a, 3)
   ElseIf D2a Like "*.?" Then
    D2a = Left(D2a, Len(D2a) - 2) & Format(FileDateTime(D2a), " yyyymmdd hhmmss") & Right(D2a, 2)
   ElseIf D2a Like "*." Then
    D2a = Left(D2a, Len(D2a) - 1) & Format(FileDateTime(D2a), " yyyymmdd hhmmss") & Right(D2a, 1)
   Else
    D2a = Left(D2a, Len(D2a) - 0) & Format(FileDateTime(D2a), " yyyymmdd hhmmss")
   End If
  End If
 End If
 
 Call doAusgabe(Ausgabe & D1 & " -> " & D2a)
 If FSO.FileExists(D2a) Then
  Dim D2an$
  D2an = REPLACE(REPLACE(LCase(D2a), ".jpg", " vorher.jpg"), ".wav", " vorher.wav")
  If FSO.FileExists(D2an) Then
   If FSO.GetFile(D2an).Size >= FSO.GetFile(D2a).Size Then ' filelen(d2an) >= filelen(d2a)
    Kill D2a
   Else
    Kill D2an
    FSO.MoveFile D2a, D2an ' Name D2a As D2an
   End If
  Else
   Call FSO.MoveFile(D2a, D2an)
  End If
 End If
 Call FSO.MoveFile(D1, D2a)
 VerschiebeFD = D2a
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in VerschiebeFD/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' VerschiebeFD

' kommt von in doAlleÜbertragen
' Wichtige IrfanView-Einstellungen:
' kein automatisches Rotieren nach Exif-Infos
' Vollbildmodus mit $D $F $X, $T $S, $M, Hintergrund rot
Function Komprimiere&(D1$, D2$, Hel&, Kontr&, Gam&, KG$, obVar%)
 On Error GoTo fehler
 Call doAusgabe("Komprimiere: ")
 If obVar Then
  Komprimiere = Shell("""" & IViewPfad & """ """ & D1 & """ /convert=""" & D2 & """ /jpgq=" & KG & " /hide=0") ' /one /fs
 Else
'   Komprimiere = Shell(vns & IViewPfad & " """ & D1 & """ /convert=""" & D2 & """ /jpgq=" & KG & " " & IIf(Hel <> 90, " /bright=" & Hel, vNS) & IIf(Kontr <> 90, " /contrast=" & Kontr, vNS)) ' /one /fs /hide=0
  Komprimiere = Shell("""" & IViewPfad & """ """ & D1 & """ /convert=""" & D2 & """ /jpgq=" & KG & " " & IIf(Hel <> 90, " /bright=" & Hel, vNS) & IIf(Kontr <> 90, " /contrast=" & Kontr, vNS)) ' /one /fs /hide=0
'  Komprimiere = Shell("""" & IViewPfad & " '" & D1 & "' /convert='" & D2 & "' /jpgq=" & KG & " " & IIf(Hel <> 90, " /bright=" & Hel, vNS) & IIf(Kontr <> 90, " /contrast=" & Kontr, vNS) & """") ' /one /fs /hide=0
 End If
 Call doAusgabe(D1 & " -> " & D2 & ", " & CStr(FSO.GetFile(D1).Size) & " -> ")
 Call WartAufProzeß(Komprimiere)
 If FSO.FileExists(D2) Then
  Call doAusgabe(CStr(FSO.GetFile(D2).Size))
 End If
 Call doAusgabe(vbCrLf)
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Komprimiere/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Komprimiere

'Public Property Let obAngezeigt(ByVal vData As Boolean)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.obAngezeigt = 5
'    mvobAngezeigt = vData
'End Property


'Public Property Get obAngezeigt() As Boolean
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.obAngezeigt
'    obAngezeigt = mvobAngezeigt
'End Property

Public Property Let PatName(ByVal vData As String)
    mvPatName = vData
End Property ' LetPatName

Public Property Get PatName() As String
    PatName = mvPatName
End Property ' GetPatName

Public Property Let Pat_id(ByVal vData As Long)
    mvPat_ID = vData
End Property

Public Property Get Pat_id() As Long
    Pat_id = mvPat_ID
End Property

Public Property Let WA(ByVal vData As String)
    mvWA = vData
End Property

Public Property Get WA() As String
    WA = mvWA
End Property

Public Property Let PatDatum(ByVal vData As Date)
    mvPatDatum = vData
End Property

Public Property Get PatDatum() As Date
    PatDatum = mvPatDatum
End Property

Public Property Let Beschreibung(ByVal vData As Variant)
    mvBeschreibung = vData
End Property

Public Property Set Beschreibung(ByVal vData As Variant)
    Set mvBeschreibung = vData
End Property

Public Property Get Beschreibung() As Variant
    If IsObject(mvBeschreibung) Then
        Set Beschreibung = mvBeschreibung
    Else
        Beschreibung = mvBeschreibung
    End If
End Property

Public Property Let KörperTeil(ByVal vData As String)
    mvKörperteil = vData
End Property

Public Property Get KörperTeil() As String
    KörperTeil = mvKörperteil
End Property

Public Property Let Pict(ByVal vData As Integer)
    mvPict = vData
End Property

Public Property Get Pict() As Integer
    Pict = mvPict
End Property

Public Property Let WavGelöscht(ByVal vData As Boolean)
    mvWavGelöscht = vData
End Property

Public Property Get WavGelöscht() As Boolean
    WavGelöscht = mvWavGelöscht
End Property

Public Property Let WavGröße(ByVal vData As Long)
    mvWavGröße = vData
End Property

Public Property Get WavGröße() As Long
    WavGröße = mvWavGröße
End Property

Public Property Let WavErstellt(ByVal vData As Date)
    mvWavErstellt = vData
End Property

Public Property Get WavErstellt() As Date
    WavErstellt = mvWavErstellt
End Property

Public Property Let wavPfad(ByVal vData As String)
    mvWavPfad = vData
End Property

Public Property Get wavPfad() As String
    wavPfad = mvWavPfad
End Property

Public Property Let NNgelöscht(ByVal vData As Boolean)
    mvNNgelöscht = vData
End Property

Public Property Get NNgelöscht() As Boolean
    NNgelöscht = mvNNgelöscht
End Property

Public Property Let NeuerName(ByVal vData As String)
    mvNeuerName = vData
End Property

Public Property Get NeuerName() As String
    NeuerName = mvNeuerName
End Property

Public Property Let verwendet(ByVal vData As Boolean)
    mvVerwendet = vData
End Property

Public Property Get verwendet() As Boolean
    verwendet = mvVerwendet
End Property

Public Property Let bearbeitet(ByVal vData As Boolean)
    mvBearbeitet = vData
End Property

Public Property Get bearbeitet() As Boolean
    bearbeitet = mvBearbeitet
End Property

Public Property Let gelöscht(ByVal vData As Boolean)
    mvGelöscht = vData
End Property

Public Property Get gelöscht() As Boolean
    gelöscht = mvGelöscht
End Property

Public Property Let bm(ByVal vData As String)
Attribute bm.VB_Description = "Bookmark auf rSteu"
'    mvBM = vData
    altPfad = vData
End Property

Public Property Get bm() As String
'    bm = mvBM
     bm = altPfad
End Property

Public Property Set Fil(ByVal vData As File)
Attribute Fil.VB_Description = "File für FileSystemObject"
    Set mvFil = vData
End Property

Public Property Get Fil() As File
    Set Fil = mvFil
End Property

Public Property Let Helligkeit(ByVal vData As Long)
    mvHelligkeit = vData
End Property

Public Property Get Helligkeit() As Long
    Helligkeit = mvHelligkeit
End Property

Public Property Let Kontrast(ByVal vData As Long)
    mvKontrast = vData
End Property

Public Property Get Kontrast() As Long
    Kontrast = mvKontrast
End Property

Public Property Let Gamma(ByVal vData As Long)
    mvGamma = vData
End Property

Public Property Get Gamma() As Long
    Gamma = mvGamma
End Property

