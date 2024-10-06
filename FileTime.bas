Attribute VB_Name = "FileTime"
Private Declare Function OpenFile Lib "kernel32" ( _
  ByVal lpFileName As String, _
  ByRef lpReOpenBuff As OFSTRUCT, _
  ByVal wStyle As Long _
  ) As Long
Private Declare Function CloseHandle _
  Lib "kernel32" ( _
  ByVal hObject As Long _
  ) As Long
  
' As Any-Deklaration von SetFileTime (fuer gewˆhnlich
' FILETIME-Parameter, die aber zu NULL gesetzt werden kˆnnen,
' um die entsprechende Information zu ignorieren
Private Declare Function SetFileTimeAPI _
  Lib "kernel32" Alias "SetFileTime" ( _
  ByVal hFile As Long, _
  ByRef lpCreationTime As Any, _
  ByRef lpLastAccessTime As Any, _
  ByRef lpLastWriteTime As Any _
  ) As Long
' Gegenst¸ck:
Private Declare Function GetFileTimeAPI _
  Lib "kernel32" Alias "GetFileTime" ( _
  ByVal hFile As Long, _
  ByRef lpCreationTime As Any, _
  ByRef lpLastAccessTime As Any, _
  ByRef lpLastWriteTime As Any _
  ) As Long
  
' Lokale Dateizeit in eine universale Dateizeit ¸bersetzen
Private Declare Function LocalFileTimeToFileTime _
  Lib "kernel32" ( _
  ByRef lpLocalFileTime As FILETIME, _
  ByRef lpFileTime As FILETIME _
  ) As Long
' Gegenst¸ck:
Private Declare Function FileTimeToLocalFileTime _
  Lib "kernel32" ( _
  ByRef lpFileTime As FILETIME, _
  ByRef lpLocalFileTime As FILETIME _
  ) As Long
  
' Eine SYSTEMTIME-Struktur in eine FILETIME ¸bersetzen
Private Declare Function SystemTimeToFileTime _
  Lib "kernel32" ( _
  ByRef lpSystemTime As SYSTEMTIME, _
  ByRef lpFileTime As FILETIME _
  ) As Long
' Gegenst¸ck:
Private Declare Function FileTimeToSystemTime _
  Lib "kernel32" ( _
  ByRef lpFileTime As FILETIME, _
  ByRef lpSystemTime As SYSTEMTIME _
) As Long
  
Private Const OF_READ = &H0
Private Const OF_READWRITE = &H2
Private Const OFS_MAXPATHNAME = 128
  
Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type
  
Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type
  
Private Type OFSTRUCT
  cBytes As Byte
  fFixedDisk As Byte
  nErrCode As Integer
  Reserved1 As Integer
  Reserved2 As Integer
  szPathName(OFS_MAXPATHNAME) As Byte
End Type
  
Public Enum FileTimeEnum
  mftCreationTime = 1
  mftLastAccessTime = 2
  mftLastWriteTime = 4
End Enum
  
  
' --- Code
  
  
Public Function GetFileTime(ByVal Pfad As String, _
                            ByVal TimeToGet As FileTimeEnum _
                           ) As Date
' Ermittelt einen der drei Zeitstempel einer Datei/eines Verzeichnisses
' und gibt diesen als Visual Basic Date-Variable zur¸ck.
Dim FTCreationTime As FILETIME, SysTime As SYSTEMTIME
Dim FTLastAccessTime As FILETIME
Dim FTLastWriteTime As FILETIME
Dim SelectedTime As FILETIME
Dim OFS As OFSTRUCT, hFile As Long
  
  ' Versuchen, die betroffene Datei zu ˆffnen
  hFile = OpenFile(Pfad, OFS, OF_READ)
  If hFile = 0 Then Exit Function ' OpenFile ist gescheitert => Ausgang
  ' Ermitteln der Zeitstempel
  GetFileTimeAPI hFile, FTCreationTime, FTLastAccessTime, FTLastWriteTime
  CloseHandle hFile
  
  ' Gesuchten Zeitstempel ausw‰hlen
  Select Case TimeToGet
    Case mftCreationTime: SelectedTime = FTCreationTime
    Case mftLastAccessTime: SelectedTime = FTLastAccessTime
    Case mftLastWriteTime: SelectedTime = FTLastWriteTime
  End Select
  
  ' Umsetzung in lokale Systemzeit
  FileTimeToLocalFileTime SelectedTime, SelectedTime
  FileTimeToSystemTime SelectedTime, SysTime
  
  ' R¸ckgabe als VB-Date
  With SysTime
    GetFileTime = _
      DateSerial(.wYear, .wMonth, .wDay) + _
      TimeSerial(.wHour, .wMinute, .wSecond)
  End With
  
End Function
  
  
Public Function SetFileTimeByDate(ByVal Pfad As String, _
                                  ByVal TimeToModify As FileTimeEnum, _
                                  ByVal DateToSet As Date)
' Setzt den Zeitstempel einer Datei unter Zuhilfenahme
' der ausf¸hrenden Funktion SetFileTime.
  
  SetFileTimeByDate = SetFileTime(Pfad, TimeToModify, _
                                  Day(DateToSet), Month(DateToSet), Year(DateToSet), _
                                  Hour(DateToSet), Minute(DateToSet), Second(DateToSet))
  
End Function
  
  
Private Function SetFileTime(ByVal Pfad As String, _
                             ByVal TimeToModify As FileTimeEnum, _
                             ByVal Tag As Integer, _
                             ByVal Monat As Integer, _
                             ByVal Jahr As Integer, _
                             ByVal Stunde As Integer, _
                             ByVal Minute As Integer, _
                             ByVal Sekunde As Integer _
                             ) As Boolean ' True => Erfolg
  
Dim FT As FILETIME, ST As SYSTEMTIME
Dim OFS As OFSTRUCT, hFile As Long, RetVal As Long
  
  ' Dateizeiten (FILETIME) und Systemzeiten (SYSTEMTIME)
  ' unterscheiden sich im Format. Zunaechst wird eine
  ' SYSTEMTIME-Struktur mit den uebergebenen Parametern gef¸llt,
  ' danach wird sie in eine FILETIME konvertiert. Da Dateizeiten
  ' GMT-orientiert geschrieben werden, ist danach noch eine Anpassung
  ' an die GMT-Zeitzone erforderlich.
  
  ' SYSTEMTIME-Struktur ausf¸llen
  With ST
    .wYear = Jahr
    .wMonth = Monat
    .wDay = Tag
    .wHour = Stunde
    .wMinute = Minute
    .wSecond = Sekunde
    '.wMilliseconds = 0
  End With
  
  ' Lokale Systemzeit in lokale Dateizeit konvertieren
  RetVal = SystemTimeToFileTime(ST, FT)
  If RetVal = 0 Then Exit Function  ' Pech gehabt
  
  ' Lokale Dateizeit in GMT-Dateizeit konvertieren
  RetVal = LocalFileTimeToFileTime(FT, FT)
  If RetVal = 0 Then Exit Function  ' Pech gehabt
  
  ' Datei fuer Lese- und Schreibzugriff ˆffnen
  hFile = OpenFile(Pfad, OFS, OF_READWRITE)
  If hFile = 0 Then Exit Function ' Pech gehabt
  
  ' Eine Datei hat 3 Dateizeiten:
  ' Erzeugung, Letzter Zugriff, Letzte Speicherung.
  ' Eine Zeit davon soll ge‰ndert werden, der Rest soll
  ' identisch bleiben. Dank "As Any"-Deklaration werden
  ' nicht zu ‰ndernde Zeitparameter mit "ByVal 0&" (f¸r
  ' NULL) bedient.
  
  If (TimeToModify And mftCreationTime) > 0 Then
    RetVal = SetFileTimeAPI(hFile, FT, ByVal 0&, ByVal 0&)
    If RetVal = 0 Then CloseHandle hFile: Exit Function ' Pech gehabt
  End If
  
  If (TimeToModify And mftLastAccessTime) > 0 Then
    RetVal = SetFileTimeAPI(hFile, ByVal 0&, FT, ByVal 0&)
    If RetVal = 0 Then CloseHandle hFile: Exit Function ' Pech gehabt
  End If
  
  If (TimeToModify And mftLastWriteTime) > 0 Then
    RetVal = SetFileTimeAPI(hFile, ByVal 0&, ByVal 0&, FT)
    If RetVal = 0 Then CloseHandle hFile: Exit Function ' Pech gehabt
  End If
  
  ' Handle schlieﬂen
  CloseHandle hFile
  
  SetFileTime = True
  
End Function


