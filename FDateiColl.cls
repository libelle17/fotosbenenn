VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FDateiColl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Auflistung aller Fotodateien"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"FDatei"
Attribute VB_Ext_KEY = "Member0" ,"FDatei"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"SteuerDB"
'Lokale Variable für eine Auflistung.
Private mCol As Collection
Private mvIndDat&
Private sCa&, sCe& ' 0 = Archiv, 1 = Eingelesene
Public fS&
Public fSGes&
Public Quelle$
Public Event indvorWechsel()
Public Event indnachWechsel()
Public Event getQuelle()
Public Event fortSchritt()

Public Function SuchItem%(si As File, Ind%)
  On Error GoTo fehler
  Call sCInit(Ind)
  ' See where the item belongs.
  Call Eingrenz(si)
  If mCol(sCa) = si Then SuchItem = -1
  Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in SuchItem/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' SuchItem
Public Function Eingrenz(new_item As File)  ', ByRef A&, ByRef E&)
 Dim Mitte&, Diff&, i&
 On Error GoTo fehler
 Diff = sCe - sCa
 If Diff < 6 Then
  For i = sCa To sCe
    If Not mCol(i).istKleinerAls(new_item) Then Exit For
  Next i
  sCa = i
  Exit Function
 End If
 Mitte = sCa + 0.5 * Diff
 If mCol(Mitte).istKleinerAls(new_item) Then
  sCa = Mitte
 Else
  sCe = Mitte
 End If
 Call Eingrenz(new_item)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Eingrenz/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Eingrenz
Public Function sCInit()
 On Error GoTo fehler
 sCa = 1
 If mCol Is Nothing Then
  Set mCol = New Collection
 Else
  sCe = mCol.Count
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in sCInit/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' sCInit

Public Sub sCAdd(ByVal new_item As File, Optional obDistinct)
  Dim i&
  On Error GoTo fehler
  If IsMissing(obDistinct) Then obDistinct = 0
  Call sCInit
  ' See where the item belongs.
  Call Eingrenz(new_item)
  ' Insert the item.
  If sCa > mCol.Count Then
   ' Add at the end.
   Add new_item
  Else
   ' Add at the right position.
   If Not obDistinct Or mCol(sCa).unGleich(new_item) Then
    Call Add(new_item, , sCa)
   End If
  End If
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in SCAdd/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub 'AddIdem

Private Function vorWechsel()
 RaiseEvent indvorWechsel ' Übertragen vom Formular in die Klassenvariable
 If mvIndDat <= 0 Then mvIndDat = 0
 If mvIndDat > mCol.Count Then mvIndDat = mCol.Count
 Call Me(mvIndDat).SchreibDatenSatz ' Übertragen von der Klassenvariable in die Datenbank
End Function
Private Function nachWechsel()
 RaiseEvent indnachWechsel ' Übertragen von der Klassenvariablen ins Formular
End Function
Public Property Let indDat(ByVal vData&)
  If indDat > 0 Then Call vorWechsel
  mvIndDat = vData
  If vData > 0 Then Call nachWechsel
End Property

Public Property Get indDat&()
  indDat = mvIndDat
End Property
Public Function Abbrechen()
  indDat = -1
End Function
Function ersterOffener&()
 Dim i&
 If rSteu Is Nothing Then
'  MsgBox "rsteu is nothing"
  Exit Function
 End If
 If rSteu.State = 0 Then
'  MsgBox "rsteu.state = 0"
  Exit Function
 End If
 If rSteu.BOF Then
'  MsgBox "rsteu.bof"
  Exit Function
 End If
 On Error GoTo f0
 For i = 1 To mCol.Count
  On Error GoTo fehler
  If mCol(i).findeSatz Then
   If IsNull(rSteu!PatName) Then GoTo gefunden
   If rSteu!PatName = vNS Then GoTo gefunden
   If rSteu!verwendet Then
    If IsNull(rSteu!Beschreibung) Then GoTo gefunden
    If rSteu!Beschreibung = vNS Then GoTo gefunden
    If IsNull(rSteu!Koerperteil) Then GoTo gefunden
    If rSteu!Koerperteil = vNS Then GoTo gefunden
    If IsNull(rSteu!WA) Then GoTo gefunden
    If rSteu!WA = vNS Then GoTo gefunden
   End If
   GoTo weiter
  Else
   GoTo weiter
  End If
gefunden:
  ersterOffener = i
  Exit Function
weiter:
 Next i
 ersterOffener = mCol.Count  ' kein Offener
 Exit Function
f0:
 If Err.Number = 9 Then
  ersterOffener = 0 ' keine Dateien
  Exit Function
 End If
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ersterOffener/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ersterOffener

Public Function Add(Fil As File, Optional sKey As String, Optional Before&) As FDatei
  Dim j&
  On Error GoTo fehler
    'neues Objekt erstellen
    Dim objNewMember As FDatei
    Set objNewMember = New FDatei
    'die Eigenschaften festlegen, die der Methode übergeben werden
    If IsObject(Fil) Then
        Set objNewMember.Fil = Fil
    Else
        Set objNewMember.Fil = Fil
    End If
    Call objNewMember.LadDatenSatz
    If Len(sKey) = 0 Then
      If Before = 0 Then
        mCol.Add objNewMember
      Else
        mCol.Add objNewMember, , Before
      End If
    Else
      mCol.Add objNewMember, sKey
    End If
    'das erstellte Objekt zurückgeben
    Set Add = objNewMember
    Set objNewMember = Nothing
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Add/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Add

Public Property Get Item(vntIndexKey As Variant) As FDatei
Attribute Item.VB_UserMemId = 0
  Dim i&, obPfad%
    'Wird verwendet, wenn auf ein Element der Auflistung zugegriffen
    'wird. vntIndexKey enthält entweder den Index oder den Schlüssel.
    'Daher wird der Parameter als Variant deklariert.
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  If VarType(vntIndexKey) = vbString Then
   obPfad = IIf(InStr(vntIndexKey, "\") > 0 Or InStr(vntIndexKey, "/") > 0, -1, 0)
   For i = 1 To mCol.Count
    If obPfad Then
     If mCol(i).mvFil.Path = vntIndexKey Then
      Set Item = mCol(i)
      Exit Property
     End If
    Else
     If mCol(i).mvFil.name = vntIndexKey Then
      Set Item = mCol(i)
      Exit Property
     End If
    End If
   Next i
  Else
   On Error Resume Next
   Set Item = mCol(vntIndexKey)
  End If
End Property

Public Property Get Count() As Long
    'Wird verwendet, wenn die Anzahl der Elemente
    'in der Auflistung ermittelt werden soll.
    'Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'Wird verwendet, wenn ein Element aus der Auflistung
    'entfernt werden soll. vntIndexKey enthält entweder
    'den Index oder den Schlüssel. Daher wird der Parameter
    'als Variant deklariert.
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'Diese Eigenschaft ermöglicht, die Elemente der
    'Auflistung mit For...Each zu durchlaufen.
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Sub Init()
 'Erstellt die Auflistung, wenn die Klasse erstellt wird.
  Dim Fil As File, Fls As Files
  Dim D2a$, DArt%
  Dim neuFDatei As FDatei
  On Error GoTo fehler
  Set mCol = New Collection
  fS = 0
  fSGes = 0
  RaiseEvent fortSchritt
  If FSO Is Nothing Then Set FSO = New FileSystemObject
  RaiseEvent getQuelle
  Set Fls = FSO.GetFolder(Quelle).Files
  fSGes = 0
  For Each Fil In Fls
   If Not Fil Like "*.modd" Then
    DArt = DateiArt(Fil.name) ' If LCase(Fil.Name) Like "*bild*.jpg" Or LCase(Fil.Name) Like "*pict*.jpg" Or LCase(Fil.Name) Like "*img_*.jpg" or lcase(fil.name) like "sam_*.jpg" Or LCase(Fil.Name) Like "*bild*.wav" Or LCase(Fil.Name) Like "*pict*.wav" Or LCase(Fil.Name) Like "*snd_*.wav" Then
    If DArt = 1 Or DArt = 2 Then
'25.5.08
     If Not Fil.name Like "*######## ######.*" Then
      D2a = Left(Fil.Path, Len(Fil.Path) - 4) & Format(FileDateTime(Fil.Path), " yyyymmdd hhmmss") & Right(Fil.Path, 4)
      Name Fil.Path As D2a
     End If
     If DArt = 1 Then fSGes = fSGes + 1
    End If
   End If
  Next Fil
'  fSGes = Fls.Count
  RaiseEvent fortSchritt
  For Each Fil In Fls
   If Not Fil Like "*.modd" Then
    If DateiArt(Fil.name) = 1 Then ' LCase(Fil.Name) Like "bild*.jpg" Or LCase(Fil.Name) Like "pict*.jpg" Then
'    Set neuFDatei = Add(Fil)
     Call sCAdd(Fil)
     fS = fS + 1
     RaiseEvent fortSchritt
     DoEvents
    End If
   End If
  Next Fil
 Exit Sub
fehler:
ErrDescription = Err.Description
ErrNumber = Err.Number
If InStrB(ErrDescription, "existiert bereits") > 0 Then
 Kill D2a
 Resume
End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Init/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' INIT


Private Sub Class_Terminate()
  Set mvarSteuerDB = Nothing
    'Entfernt die Auflistung aus dem Speicher, wenn die Klasse beendet wird.
    Set mCol = Nothing
End Sub

