Attribute VB_Name = "Speichern"
' Dieser Source stammt von http://www.activevb.de
' und kann frei verwendet werden. Für eventuelle Schäden
' wird nicht gehaftet.
'
' Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
' Ansonsten viel Spaß und Erfolg mit diesem Source !
'
' Autor: K. Langbein Klaus@ActiveVB.de
'
' Beschreibung:
' VB kann Bitmaps immer nur im Format der aktuellen
' Monitoreinstellung abspeichern. Bei eingachen Grafiken ist es
' jedoch oft sinnvol ein Bild mit geringerer Farbauflösung oder
' sogar in schwarz/weiß abszuspeichern. Mit Hilfe des API-Befehls
' GetDiBits können geräteunabhängige Bitmaps (DIBs) aus dem Image
' oder Picture einer Picturebox erstellt werden. Das resultierende
' Datenfeld kann dann zusammen mit dem entsprechenden Header als
' DIB mit der gewünschten Farbtiefe abgespeichert werden.
'
' Chronologie und Referenzen
' Ein ähnliches Programm wird in den Beispielen zur
' Programmiersprache GFA-Basic (www.gfasoft.gfa.net) mitgeliefert.
' Dieser Sourcecode wurde mit Mühen und etlichen Abstürzen nach VB
' portiert und optimiert. Erst durch einen Hinweis in "VB Programmer's
' Guide to the Win32 API" von Dan Appleman wurde die Abfrage der
' resultierenden Dateigröße möglich: Der Wert 0 als Stellvertreter
' für die Übergabe des Datenfeldes muß mit Byval übergeben werden!
'
' Der Anstoß zur Erstellung dieses Tipps kam durch wiederholte
' Anfragen im ActiveVB-Forum zustande. Die letzte Anfrage kam von
' Siml <blue-siml@gmx.net>. Er hat sich daher bereit erklärt die
' Benutzeroberfläche etwas auszubauen und das Testbild und einen
' Teil der Kommentierung hinzuzufügen.

Option Explicit
Public Const DIB_RGB_COLORS = 0

Public Declare Function GetObjectgdi32 Lib "GDI32" Alias _
        "GetObjectA" (ByVal hObject As Long, ByVal nCount As _
        Long, lpObject As Any) As Long

' Dies ist die übliche Deklaration für GetDIBits. Sie wird hier
' in dieser Form nicht verwendet, da die Struktur BITMAPINFO
' in modifizierter Form übergeben werden muß.
Public Declare Function GetDIBits Lib "GDI32" (ByVal _
        aHDC As Long, ByVal hBitmap As Long, ByVal _
        nStartScan As Long, ByVal nNumScans As Long, lpBits _
        As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) _
        As Long

' Deklaration für Übergabe vo BITMAPINFO256, welche Platz für eine
' 256 Byte lange Farbpalette enthält.
Public Declare Function GetDIBits256 Lib "GDI32" Alias "GetDIBits" _
       (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal _
        nStartScan As Long, ByVal nNumScans As Long, lpBits _
        As Any, lpBI As BITMAPINFO256, ByVal wUsage As Long) _
        As Long

Public Type BITMAPINFOHEADER
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

Public Type BITMAP
  bmType As Long
  BmWidth As Long
  BmHeight As Long
  bmWidthBytes As Long
  BmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

' Die Farben einer Bitmap sind üblicherweise so wie im unten
' definierten Typ RGBQUAD angeordnet. D.h. Blau steht an erster,
' und Rot steht an dritter Stelle. In einer Variablen vom Typ
' Long steht der Rotwert im niederwertigsten Byte, was jedoch
' auf Grund der Vertauschung der Bytefolge bei Intelprozessoren
' im Speicher links steht. RGBQUAD wird hier nicht verwendet;
' stattdessen wird die Funktion SwapRedBlue eingesetzt, um die
' richtige Bytefolge für Farbpaletten zu erhalten .
Public Type RGBQUAD
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
  rgbReserved As Byte
End Type

' Dies ist die übliche Deklaration für BITMAPINFO -
' sie wird hier nicht verwendet.
Public Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As Long
End Type

' Dies ist die Stuktur die hier zum Einsatz kommt. Sie
' ist groß genug um eine Farbpalette mit einer Länge von 256
' Byte aufzunehmen. Übergibt man beim Aufruf von GetDiBits die
' oben deklarierte Datenstruktur, kommt es zum Absturz, da der
' Speicherbereich oberhalb von BITMAPINFO mit den Daten der
' Palette überschrieben wird.
Public Type BITMAPINFO256
  bmiHeader As BITMAPINFOHEADER
  bmiColors(255) As Long
End Type



Function SaveBitmap_AllRes(hDC As Long, handle As Long, _
                           ByVal BitsPerPixel As Long, _
                           ByVal FName$, _
                           Optional NewPal As Variant) As Long
    On Error GoTo err1
    
    Dim bmp As BITMAP
    Dim i As Integer
    Dim bInfo As BITMAPINFO256
    Dim FileHeader As BITMAPFILEHEADER
    Dim bArray() As Byte
    Dim nLines As Long
    Dim WidthArray As Long
    Dim fno As Long
    Dim Palette() As Long
    Dim newP As Long
    Dim nCol As Long
    
    Call GetObjectgdi32(handle, Len(bmp), bmp)
    
    Select Case BitsPerPixel
    Case 1, 4, 8, 16, 24, 32
        ' kein Fehler. weiter gehts!
    Case Else
        MsgBox "Fehler!" & vbCrLf & "Dieses Bildformat wird nicht unterstützt!"
        SaveBitmap_AllRes = -1
        Exit Function
    End Select
   
    If IsMissing(NewPal) = 0 Then
        newP = UBound(NewPal)
    End If
     
    bInfo.bmiHeader.biHeight = bmp.BmHeight
    bInfo.bmiHeader.biWidth = bmp.BmWidth
    bInfo.bmiHeader.biPlanes = bmp.BmPlanes
    bInfo.bmiHeader.biBitCount = BitsPerPixel
    bInfo.bmiHeader.biSize = Len(bInfo.bmiHeader)
    bInfo.bmiHeader.biCompression = 0
    
    ' Der 1. Aufruf ohne Übergabe von bArray, dient dazu die Größe
    ' des benötigten Feldes festzustellen. Die Palette wir hier auch
    ' schon übertragen.
    nLines = GetDIBits256(hDC, handle, 0, bmp.BmHeight, _
                       ByVal 0, bInfo, DIB_RGB_COLORS)
    
    If nLines = 0 Then         ' Falls ein Fehler auftrat wird nLines 0,
        SaveBitmap_AllRes = -2 ' sonst ist es die Zahl der Zeilen.
        Exit Function
    End If
                       
    ' Jetzt können wir die Breite einer Zeile berechnen. Diese ist
    ' nicht notwendigerweise wie erwartet, sondern enthält evtl.
    ' sogenannte Padbytes.
    WidthArray = bInfo.bmiHeader.biSizeImage / bInfo.bmiHeader.biHeight
    ReDim bArray(1 To WidthArray, 1 To bInfo.bmiHeader.biHeight)
    
    ' Jetzt wird tatsächlich gelesen. Die Bitmapdaten befinden sich
    ' anschließend in bArray und könnten hier auch manipuliert werden.
    nLines = GetDIBits256(hDC, handle, 0, bmp.BmHeight, _
                        bArray(1, 1), bInfo, DIB_RGB_COLORS)
                        
    If nLines = 0 Then
        SaveBitmap_AllRes = -3 ' Tja, dann ist wohl was schiefgelaufen.
        Exit Function
    End If
    
    Select Case BitsPerPixel

    Case 1
        bInfo.bmiHeader.biClrUsed = 2
        bInfo.bmiHeader.biClrImportant = 2
        nCol = 1
    Case 4
        bInfo.bmiHeader.biClrUsed = 16
        bInfo.bmiHeader.biClrImportant = 16
        nCol = 15
    Case 8
        bInfo.bmiHeader.biClrUsed = 256
        bInfo.bmiHeader.biClrImportant = 256
        nCol = 255
    Case 16, 24, 32
        nCol = 0
    End Select
    
    ReDim Palette(nCol)
    ' Hier wird umgespeichert, damit wir die Palette einfach mit Put
    ' ausgeben können. Gleichzeitig können wir eine an die Funktion
    ' übergebene Palette (NewPal) verwenden.
    If nCol > 0 Then
        If newP = nCol Then
            For i = 0 To nCol
                Palette(i) = SwapRedBlue(NewPal(i)) ' Rot und Blau sind
                ' in normalen Longs vertauscht. Wir korrigieren das hier.
            Next i
        Else
            For i = 0 To nCol
                Palette(i) = bInfo.bmiColors(i)
            Next i
        End If
    Else
        Palette(0) = 0
    End If
    
    FileHeader.bfType = 19778 ' entspricht "BM"
    FileHeader.bfOffBits = Len(FileHeader) + Len(bInfo.bmiHeader)
    FileHeader.bfOffBits = FileHeader.bfOffBits + (UBound(Palette) + 1) * 4
    FileHeader.bfSize = Len(FileHeader) + Len(bInfo.bmiHeader) + (UBound(Palette) + 1) * 4
    
    fno = FreeFile
    Open FName$ For Binary As #fno
    ' und wieder Ausspucken...
    Put #fno, , FileHeader
    Put #fno, , bInfo.bmiHeader
    Put #fno, , Palette()
    Put #fno, , bArray()
    Close #fno
    SaveBitmap_AllRes = FileLen(FName$)
    
    Exit Function
    
err1:
    Select Case Err
    
    Case 999
    
    Case Else
        MsgBox "Fehler!" & vbCrLf & Error$
        'Resume
    End Select
    
End Function
Function SwapRedBlue(ByVal Col As Long) As Long
    
    Dim red As Long, green As Long, blue As Long, newcolor As Long
    red = Col And 255
    green = Col And 65280
    green = green / 256   '2 ^ 8
    blue = Col And 16711680
    blue = blue / 65536 ' 2^16
            
    newcolor = blue + green * 256 + red * 65536
    SwapRedBlue = newcolor
    
End Function



