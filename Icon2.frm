VERSION 5.00
Begin VB.Form fürIcon 
   Caption         =   "Fotos benennen"
   ClientHeight    =   13680
   ClientLeft      =   2415
   ClientTop       =   1635
   ClientWidth     =   14790
   Icon            =   "Icon2.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Icon2.frx":1601A
   ScaleHeight     =   13680
   ScaleWidth      =   14790
   Begin VB.CommandButton nädP 
      Caption         =   "nä&.d.Pat."
      Height          =   255
      Left            =   12600
      TabIndex        =   82
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton rueckgaengig 
      Caption         =   "&rückgängig"
      Height          =   255
      Left            =   10800
      TabIndex        =   55
      Top             =   9840
      Width           =   1455
   End
   Begin VB.TextBox Kompressionsgrad 
      Height          =   285
      Left            =   13440
      TabIndex        =   51
      Top             =   9720
      Width           =   375
   End
   Begin VB.CommandButton nur2 
      Caption         =   "nur &2 Wö"
      Height          =   255
      Left            =   10200
      TabIndex        =   81
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton ldP 
      Caption         =   "&ldP"
      Height          =   315
      Left            =   12480
      TabIndex        =   80
      Top             =   4350
      Width           =   495
   End
   Begin VB.CommandButton ndP 
      Caption         =   "&ndP"
      Height          =   315
      Left            =   13080
      TabIndex        =   79
      Top             =   4350
      Width           =   495
   End
   Begin VB.CommandButton wieLetztesdPat 
      Caption         =   "w&ie letztes d.Pat."
      Height          =   255
      Left            =   11040
      TabIndex        =   78
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton wieNächstes 
      Caption         =   "näch&stes"
      Height          =   255
      Left            =   12960
      TabIndex        =   77
      Top             =   5570
      Width           =   855
   End
   Begin VB.CommandButton BearbeiteteNochmalVerschieben 
      Caption         =   "B&earbeitete nochmal verschieben"
      Height          =   375
      Left            =   9000
      TabIndex        =   76
      Top             =   1270
      Width           =   2535
   End
   Begin VB.CommandButton FtCn 
      Caption         =   "&F"
      Height          =   345
      Left            =   7680
      TabIndex        =   5
      Top             =   900
      Width           =   200
   End
   Begin VB.CommandButton CnStr 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1270
      Width           =   8895
   End
   Begin VB.CommandButton RechtsNeu 
      Caption         =   "90° re neu"
      Height          =   255
      Left            =   12120
      TabIndex        =   40
      Top             =   8560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton LinksNeu 
      Caption         =   "90° li neu"
      Height          =   255
      Left            =   11160
      TabIndex        =   39
      Top             =   8560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton FreiDrehen 
      Caption         =   "° drehen"
      Height          =   255
      Left            =   10200
      TabIndex        =   38
      Top             =   8560
      Width           =   855
   End
   Begin VB.TextBox Grad 
      Height          =   285
      Left            =   9480
      TabIndex        =   37
      Top             =   8520
      Width           =   615
   End
   Begin FotosBenennen.sssKalender BildDatum 
      Height          =   255
      Left            =   9480
      TabIndex        =   74
      Top             =   6720
      Width           =   2775
      _extentx        =   4895
      _extenty        =   450
      scaleheight     =   255
      scalemode       =   0
   End
   Begin VB.TextBox GamZ 
      Height          =   285
      Left            =   12960
      TabIndex        =   45
      Top             =   8955
      Width           =   615
   End
   Begin VB.TextBox KontrZ 
      Height          =   285
      Left            =   12240
      TabIndex        =   44
      Top             =   8955
      Width           =   615
   End
   Begin VB.TextBox HelZ 
      Height          =   285
      Left            =   11520
      TabIndex        =   43
      Top             =   8955
      Width           =   615
   End
   Begin VB.CommandButton FarbenZurück 
      Caption         =   "F&arben zurück"
      Height          =   255
      Left            =   9480
      TabIndex        =   41
      Top             =   8955
      Width           =   1335
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   100
      Left            =   11400
      Max             =   500
      Min             =   1
      TabIndex        =   36
      Top             =   8430
      Value           =   250
      Width           =   2295
   End
   Begin VB.CheckBox stehenLassen 
      Caption         =   "stehenlassen"
      Height          =   195
      Left            =   10800
      TabIndex        =   48
      Top             =   9600
      Width           =   1335
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   100
      Left            =   11400
      Max             =   200
      TabIndex        =   35
      Top             =   8325
      Value           =   90
      Width           =   2295
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   100
      Left            =   11400
      Max             =   200
      TabIndex        =   34
      Top             =   8220
      Value           =   90
      Width           =   2295
   End
   Begin VB.OptionButton obMySQL 
      Caption         =   "MySQL"
      Height          =   195
      Left            =   6600
      TabIndex        =   71
      Top             =   720
      Width           =   885
   End
   Begin VB.OptionButton obAcc 
      Caption         =   "Access"
      Height          =   195
      Left            =   6600
      TabIndex        =   70
      Top             =   480
      Width           =   945
   End
   Begin VB.CheckBox keinTon 
      Caption         =   "kein Ton&?"
      Height          =   195
      Left            =   11760
      TabIndex        =   69
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton ersterOffenerCmd 
      Caption         =   "1.O&ffener"
      Height          =   315
      Left            =   11850
      TabIndex        =   68
      Top             =   4680
      Width           =   795
   End
   Begin VB.CommandButton wieVoriges 
      Caption         =   "wieVori&ges"
      Height          =   255
      Left            =   11880
      TabIndex        =   67
      Top             =   5570
      Width           =   975
   End
   Begin VB.CommandButton Lad 
      Caption         =   "Ak&tualis."
      Height          =   315
      Left            =   12600
      TabIndex        =   66
      Top             =   4680
      Width           =   795
   End
   Begin VB.TextBox Doppler 
      BackColor       =   &H80000010&
      ForeColor       =   &H00FF0000&
      Height          =   2235
      Left            =   9600
      MultiLine       =   -1  'True
      TabIndex        =   64
      Top             =   11280
      Width           =   4215
   End
   Begin VB.TextBox Fußstatus 
      BackColor       =   &H80000010&
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   9570
      MultiLine       =   -1  'True
      TabIndex        =   62
      Top             =   10320
      Width           =   4185
   End
   Begin VB.TextBox DSZahl 
      BackColor       =   &H80000010&
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   11640
      TabIndex        =   61
      Top             =   4350
      Width           =   795
   End
   Begin VB.TextBox Position 
      BackColor       =   &H80000010&
      Enabled         =   0   'False
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   10920
      TabIndex        =   60
      Top             =   4350
      Width           =   675
   End
   Begin VB.CommandButton ZumEnde 
      Caption         =   "&Z"
      Height          =   315
      Left            =   11550
      TabIndex        =   59
      Top             =   4680
      Width           =   315
   End
   Begin VB.CommandButton ZumAnfang 
      Caption         =   "&1"
      Height          =   315
      Left            =   9510
      TabIndex        =   56
      Top             =   4680
      Width           =   285
   End
   Begin VB.TextBox Schrittweite 
      Height          =   285
      Left            =   10440
      TabIndex        =   58
      Text            =   "1"
      Top             =   4350
      Width           =   375
   End
   Begin VB.TextBox NeuerName 
      BackColor       =   &H80000016&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   9510
      TabIndex        =   54
      Top             =   9270
      Width           =   4275
   End
   Begin VB.TextBox DateiZeit 
      BackColor       =   &H80000004&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd.mm.yy  hh:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1031
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   12150
      TabIndex        =   53
      Top             =   6720
      Width           =   1665
   End
   Begin VB.TextBox DateiBreite 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   285
      Left            =   5370
      TabIndex        =   52
      Top             =   480
      Width           =   825
   End
   Begin VB.TextBox DateiHöhe 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   4500
      TabIndex        =   50
      Top             =   480
      Width           =   765
   End
   Begin VB.CommandButton rechtsCmd 
      Caption         =   "9&0° re"
      Height          =   315
      Left            =   10500
      TabIndex        =   33
      Top             =   8220
      Width           =   855
   End
   Begin VB.CommandButton LinksCmd 
      Caption         =   "&90° li"
      Height          =   315
      Left            =   9510
      TabIndex        =   32
      Top             =   8220
      Width           =   915
   End
   Begin VB.CommandButton AlleÜbertragen 
      Caption         =   "alle Übertragen"
      Height          =   405
      Left            =   9510
      TabIndex        =   47
      Top             =   9630
      Width           =   1275
   End
   Begin VB.ComboBox Armstrong 
      Height          =   315
      Left            =   10290
      TabIndex        =   27
      Top             =   7880
      Width           =   3015
   End
   Begin VB.ComboBox Wagner 
      Height          =   315
      Left            =   10290
      TabIndex        =   25
      Top             =   7580
      Width           =   3015
   End
   Begin VB.ComboBox Beschreibung 
      Height          =   315
      Left            =   9510
      TabIndex        =   23
      Top             =   7260
      Width           =   4245
   End
   Begin VB.CommandButton TonCmd 
      Caption         =   "T&on"
      Height          =   315
      Left            =   13320
      TabIndex        =   17
      Top             =   4680
      Width           =   435
   End
   Begin VB.CommandButton VorwärtsCmd 
      Caption         =   "vorw&ärts"
      Height          =   315
      Left            =   10740
      TabIndex        =   16
      Top             =   4680
      Width           =   795
   End
   Begin VB.CommandButton RückwärtsCmd 
      Caption         =   "r&ückwärts"
      Height          =   315
      Left            =   9840
      TabIndex        =   15
      Top             =   4680
      Width           =   855
   End
   Begin VB.CheckBox obPat 
      Caption         =   "Bild für T&urbomed"
      Height          =   285
      Left            =   10200
      TabIndex        =   14
      Top             =   5610
      Width           =   1635
   End
   Begin VB.ComboBox KörperTeil 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9510
      TabIndex        =   21
      Top             =   6180
      Width           =   4185
   End
   Begin VB.TextBox NamDB 
      Height          =   345
      Left            =   9570
      TabIndex        =   31
      Top             =   60
      Width           =   4785
   End
   Begin VB.TextBox Kompr 
      Height          =   345
      Left            =   9570
      TabIndex        =   29
      Top             =   480
      Width           =   4785
   End
   Begin VB.ComboBox PatName 
      Height          =   315
      Left            =   9510
      TabIndex        =   19
      Top             =   5250
      Width           =   4185
   End
   Begin VB.TextBox SteuerDB 
      Height          =   345
      Left            =   4560
      TabIndex        =   4
      Top             =   900
      Width           =   3135
   End
   Begin VB.ListBox Lw 
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3075
   End
   Begin VB.TextBox Ausgabe 
      Height          =   2625
      Left            =   540
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   13
      Top             =   1650
      Width           =   12975
   End
   Begin VB.CommandButton Abbruch 
      Caption         =   "Abbru&ch"
      Height          =   375
      Left            =   3180
      TabIndex        =   12
      Top             =   450
      Width           =   1095
   End
   Begin VB.CommandButton Start 
      Caption         =   "&Start"
      Height          =   375
      Left            =   3150
      TabIndex        =   11
      Top             =   30
      Width           =   765
   End
   Begin VB.TextBox Archiv 
      Height          =   345
      Left            =   9540
      TabIndex        =   10
      Top             =   1260
      Width           =   4785
   End
   Begin VB.TextBox ArchPat 
      Height          =   375
      Left            =   9540
      TabIndex        =   8
      Top             =   870
      Width           =   4785
   End
   Begin VB.TextBox Quelle 
      Height          =   375
      Left            =   660
      TabIndex        =   2
      Top             =   870
      Width           =   3045
   End
   Begin VB.TextBox DateiPfad 
      BackColor       =   &H80000004&
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3960
      TabIndex        =   46
      Top             =   30
      Width           =   3165
   End
   Begin VB.PictureBox Picture2 
      DrawMode        =   1  'Schwarzintensität
      Height          =   6495
      Left            =   120
      ScaleHeight     =   6435
      ScaleWidth      =   7755
      TabIndex        =   72
      Top             =   4320
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.Label DTyp 
      Height          =   255
      Left            =   9600
      TabIndex        =   83
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label KompressionsgradLbl 
      Caption         =   "&Kompr'grad:"
      Height          =   255
      Left            =   12360
      TabIndex        =   49
      Top             =   9720
      Width           =   975
   End
   Begin VB.Label Version 
      Height          =   255
      Left            =   7200
      TabIndex        =   75
      Top             =   120
      Width           =   885
   End
   Begin VB.Label HellLab 
      Caption         =   "&Hell.:"
      Height          =   255
      Left            =   11040
      TabIndex        =   42
      Top             =   8955
      Width           =   375
   End
   Begin VB.Label Dopplerlabel 
      Caption         =   "Doppler vom"
      Height          =   285
      Left            =   9600
      TabIndex        =   65
      Top             =   11040
      Width           =   4455
   End
   Begin VB.Label FußstatusBez 
      Caption         =   "Fußstatus vom"
      Height          =   285
      Left            =   9600
      TabIndex        =   63
      Top             =   10110
      Width           =   4575
   End
   Begin VB.Label SchrittweiteBez 
      Caption         =   "Schrittw&eite"
      Height          =   225
      Left            =   9450
      TabIndex        =   57
      Top             =   4380
      Width           =   855
   End
   Begin VB.Label NamDBBez 
      Caption         =   "&Namensdatenbank:"
      Height          =   345
      Left            =   8160
      TabIndex        =   28
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label ArmstrongBez 
      Caption         =   "A&rmstrong:"
      Height          =   225
      Left            =   9510
      TabIndex        =   26
      Top             =   7880
      Width           =   765
   End
   Begin VB.Label WagnerBez 
      Caption         =   "&Wagner:"
      Height          =   255
      Left            =   9540
      TabIndex        =   24
      Top             =   7580
      Width           =   705
   End
   Begin VB.Label BeschreibungBez 
      Caption         =   "&Beschreibung:"
      Height          =   225
      Left            =   9540
      TabIndex        =   22
      Top             =   7020
      Width           =   1155
   End
   Begin VB.Label BildDatumBez 
      Caption         =   "Bild&-Datum"
      Height          =   225
      Left            =   9510
      TabIndex        =   73
      Top             =   6480
      Width           =   1065
   End
   Begin VB.Label KörperTeilBez 
      Caption         =   "&Körperteil"
      Height          =   195
      Left            =   9510
      TabIndex        =   20
      Top             =   5970
      Width           =   1485
   End
   Begin VB.Label PatNamBez 
      Caption         =   "&Patient oder Beschreibung:"
      Height          =   195
      Left            =   9510
      TabIndex        =   18
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label KomprBez 
      Caption         =   "&Verz.f.kompr.Dateien:"
      Height          =   285
      Left            =   8040
      TabIndex        =   30
      Top             =   540
      Width           =   1575
   End
   Begin VB.Label SteuerDBBez 
      Caption         =   "Steuer&db:"
      Height          =   315
      Left            =   3810
      TabIndex        =   3
      Top             =   930
      Width           =   705
   End
   Begin VB.Image Image1 
      Height          =   7725
      Left            =   120
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   9345
   End
   Begin VB.Image Status 
      Height          =   480
      Left            =   60
      Top             =   1650
      Width           =   480
   End
   Begin VB.Label ArchivBez 
      Caption         =   "Arc&hiv andere:"
      Height          =   255
      Left            =   7890
      TabIndex        =   9
      Top             =   1290
      Width           =   1635
   End
   Begin VB.Label ArchPatBez 
      Caption         =   "&Archiv Patientenbilder:"
      Height          =   255
      Left            =   7890
      TabIndex        =   7
      Top             =   930
      Width           =   1665
   End
   Begin VB.Label QuelleLabel 
      Caption         =   "&Quelle:"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   960
      Width           =   585
   End
End
Attribute VB_Name = "fürIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents FDC As FDateiColl
Attribute FDC.VB_VarHelpID = -1
Public WithEvents FDt As FDatei
Attribute FDt.VB_VarHelpID = -1
Public WithEvents dbv As DBVerb
Attribute dbv.VB_VarHelpID = -1
Private letzterSatz&
'Public PatInfosNichtNeu%
Private BildNichtNeu%
Private Declare Function BitBlt& Lib "GDI32" _
 (ByVal hDestDC&, ByVal x&, ByVal Y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal xSrc&, ByVal ySrc&, ByVal dwRop&)
Private Declare Function GetObjectgdi32& Lib "GDI32" Alias "GetObjectA" (ByVal hObject&, ByVal nCount&, lpObject As Any)
Private Declare Function VarPtrArray& Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen&)

Private Type SAFEARRAYBOUND
  cElements As Long
  lLbound As Long
End Type

Private Type SAFEARRAY1D
  cDims As Integer
  fFeatures As Integer
  cbElements As Long
  cLocks As Long
  pvData As Long
  Bounds(0 To 0) As SAFEARRAYBOUND
End Type

Private Type SAFEARRAY2D
  cDims As Integer
  fFeatures As Integer
  cbElements As Long
  cLocks As Long
  pvData As Long
  Bounds(0 To 1) As SAFEARRAYBOUND
End Type


Private Type BITMAP
  bmType As Long
  BmWidth As Long
  BmHeight As Long
  bmWidthBytes As Long
  BmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Private Const SRCCOPY = &HCC0020

Dim Brightness!, Contrast!, Gamma!
Dim neuBrightness!, neuContrast!, neuGamma!
Dim TableBright(255) As Byte

Private Sub rueckgaengig_Click() ' rückgängig
 Dim erg$, erg0$, rs As New Recordset, i%, zahl%
 ConstrFestleg Me, 0
 erg = Dir(Me.Kompr & "*foto*.jpg")
 Do While erg <> vNS
  erg0 = erg
  For i = 0 To 5
   rs.Open "SELECT * FROM `jpg` where concat(neuername,'.jpg') = '" & erg & "'", FotoCn, adOpenStatic, adLockReadOnly
   If Not rs.EOF Then
    Exit For
   Else
    erg = REPLACE$(erg, " WA", "  WA")
    rs.Close
   End If
  Next i
  If rs.State = 0 Then
   MsgBox erg0 & vbCrLf & " nicht in der Datenbank auffindbar."
  Else
   If FSO.FileExists(rs!wavPfad) Then
    Dim dire$
    dire = vNS
    If Not IsNull(rs!Directory) Then If LenB(rs!Directory) <> 0 Then dire = rs!Directory
    If LenB(dire) = 0 Then dire = Left(rs!Pfad, Len(rs!Pfad) - Len(rs!name))
    FSO.MoveFile rs!wavPfad, dire
   End If
   If FSO.FileExists(rs!tPfad) Then
    FSO.MoveFile rs!tPfad, rs!Pfad
    Kill Me.Kompr & erg0
    zahl = zahl + 1
   End If
   rs.Close
  End If
  erg = Dir
 Loop
 MsgBox zahl & " Dateien wieder bereitgestellt."
End Sub ' rueckgaengig_Click()
Private Sub BearbeiteteNochmalVerschieben_Click()
 Dim erg$, rs As New ADODB.Recordset
 erg = Dir(Quelle)
 If Right(Quelle, 1) <> "\" Then Quelle = Quelle & "\"
 If Right(ArchPat, 1) <> "\" Then ArchPat = ArchPat & "\"
 FotoCn.Open dbv.cnVorb(vNS, "jpg", "fotosinp")
 Do While erg <> vNS
  Set rs = Nothing
  rs.Open "SELECT * FROM `jpg` where name = '" & IIf(InStr(FotoCn, "MySQL") > 0 Or InStr(FotoCn, "MSDASQL") > 0, REPLACE(erg, "\", "\\"), erg) & "' or wavpfad like '%" & IIf(InStr(FotoCn, "MySQL") > 0 Or InStr(FotoCn, "MSDASQL") > 0, REPLACE(erg, "\", "\\"), erg) & "'", FotoCn, adOpenStatic, adLockReadOnly
  If Not rs.EOF Then
   If rs!bearbeitet And Not IsNull(rs!NeuerName) And Trim(rs!NeuerName) <> vNS Then
    On Error Resume Next
    Name Quelle & erg As ArchPat & erg
    If Err.Number <> 0 Then
     On Error GoTo fehler
     Dim fd1 As Date, fd2 As Date, fl1&, fl2&
     fd1 = FileDateTime(Quelle & erg)
     fd2 = FileDateTime(ArchPat & erg)
     fl1 = FileLen(Quelle & erg)
     fl2 = FileLen(ArchPat & erg)
     If fd1 = fd2 And fl1 = fl2 Then
      Kill Quelle & erg
     End If
    Else
     On Error GoTo fehler
    End If
    
   End If
  End If
  erg = Dir
 Loop
fehler:
End Sub

Private Sub CnStr_Click()
 Call ConstrFestleg(Me, 0)
' Call dbv.Auswahl(vNS, "anamnesebogen", "Quelle")
' QuelCStr = dbv.CnStr
' Me.CnStr.Caption = dbv.Constr
End Sub

'Private Sub dbv_wCnAendern(dbvCnStr As String)
' If obMySQL Then If dbv.Ü2 = "Quelle" Then CnStr.Caption = dbvCnStr
'End Sub ' dbv_wCnAendern(dbvCnStr As String)

Private Sub FarbenZurück_Click()
 BildNichtNeu = -1
 Me.HScroll1 = 90
 Me.HScroll2 = 90
 Me.HScroll3 = 250
 BildNichtNeu = 0
 Call HellKontr
End Sub


Private Sub FDC_fortSchritt()
 Select Case FDC.fS
  Case 0: If FDC.fSGes = 0 Then Me.DateiPfad = "..." Else Me.DateiPfad = "0 von " + CStr(FDC.fSGes)
  Case Else: Me.DateiPfad = CStr(FDC.fS) + " von " + CStr(FDC.fSGes)
 End Select
' Me.Refresh
End Sub

Private Sub FDC_getQuelle()
 FDC.Quelle = Me.Quelle
End Sub

Private Sub FDC_indnachWechsel() ' Übertragen von der Klassenvariablen ins Formular
 Dim i&, WT$
 On Error GoTo fehler
 Call BeginWarten
 Set FDt = Me.FDC(Me.FDC.indDat)
 If Not FDt Is Nothing Then
 With FDt
  Me.PatName = .PatName
  Me.obPat = -.verwendet
  Me.KörperTeil = .KörperTeil
  Me.BildDatum.Value = .PatDatum
  Me.Beschreibung = .Beschreibung
  If Not IsNull(.WA) Then
'   Me.Wagner = WagnerText(Me, .WA)
   WT = WagnerText(Me, .WA)
   Me.Wagner = WT
   For i = 0 To Me.Wagner.ListCount - 1
    If Me.Wagner.List(i) = Me.Wagner.Text Then
     Me.Wagner.ListIndex = i
     Exit For
    End If
   Next
'   Me.WagnerL.Text = WT
   Me.Armstrong = ArmstrongText(Me, .WA)
  End If
  Me.NeuerName = .NeuerName
  BildNichtNeu = -1
  Me.HScroll1 = IIf(.Helligkeit = 0, 90, .Helligkeit)
  Me.HScroll2 = IIf(.Kontrast = 0, 90, .Kontrast)
  Me.HScroll3 = IIf(.Gamma = 0, 250, .Gamma)
  BildNichtNeu = 0
  Call DateiAnzeig(Me)
  Me.RückwärtsCmd.Enabled = 0
  Me.VorwärtsCmd.Enabled = 0
  Me.ldP.Enabled = 0
  Me.ndP.Enabled = 0
  If Me.FDC.Count > -1 Then
   If Me.FDC.indDat > 1 Then Me.RückwärtsCmd.Enabled = True
   If Me.FDC.indDat < Me.FDC.Count Then Me.VorwärtsCmd.Enabled = True
   Me.ldP.Enabled = True
   Me.ndP.Enabled = True
  End If
  If letzterSatz <> FDC.indDat And Me.keinTon = False Then Call TonCmd_Click ' 1 = true
  If .WavGelöscht Then
   Me.TonCmd.Enabled = False
  Else
   Me.TonCmd.Enabled = True
  End If
 End With
 End If
 Call doPatNameChange(Me)
 Call EndeWarten
 Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in FDC_indnachWechsel/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' FDC_indnachWechsel

Public Sub BeginWarten()
' Me.Visible = False
 Dim ctl As Control
 For Each ctl In Me.Controls
  ctl.Visible = False
 Next ctl
End Sub ' BeginWarten

Public Sub EndeWarten()
' Me.Visible = True
 Dim ctl As Control
 For Each ctl In Me.Controls
  If ctl.name <> "Picture2" Then
   ctl.Visible = True
  End If
 Next ctl
End Sub ' EndeWarten

Public Sub FDC_indvorWechsel() ' Übertragen vom Formular in die Klassenvariable
' With Me.FDC(Me.FDC.indDat)
 On Error GoTo fehler
  If Not FDt Is Nothing Then
  With FDt
  .PatName = Me.PatName
  .Pat_id = getPat_id(Me.PatName)
  .verwendet = -Me.obPat
  .KörperTeil = Me.KörperTeil
  .PatDatum = Me.BildDatum.Value
  .Beschreibung = Me.Beschreibung
  If .Helligkeit <> Me.HScroll1 Or .Kontrast <> Me.HScroll2 Or .Gamma <> Me.HScroll3 Then
   .Helligkeit = Me.HScroll1
   .Kontrast = Me.HScroll2
   .Gamma = Me.HScroll3
   If Me.HScroll1.Value <> 90 Or Me.HScroll2.Value <> 90 Or Me.HScroll3.Value <> 250 Then
    Call ConvertTojpg(Me.Image1.Picture, getVariantePfad(FDt.Fil)) ' Left(FDt.Fil.Path, Len(FDt.Fil.Path) - Len(FDt.Fil.Name)) & "v" & FDt.Fil.Name
   End If
  End If
'  .WA = IIf(Len(Me.WagnerL) > 0, Left(Me.WagnerL, 1), " ") + UCase(Left(Me.Armstrong, 1))
  .WA = IIf(Len(Me.Wagner) > 0, Left(Me.Wagner, 1), " ") + UCase(Left(Me.Armstrong, 1))
  .NeuerName = Me.NeuerName
  Call neutralisier
 End With
 End If
 letzterSatz = Me.FDC.indDat
 Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in FDC_indvorWechsel/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' FDC_indvorWechsel

Private Function ConvertTojpg&(Pict As IPictureDisp, Ziel$, Optional Quality% = 100)
  Dim Path$, jpg As New cDIBSection
  On Error GoTo fehler
  jpg.CreateFromPicture Pict
  If Not Savejpg(jpg, Ziel, Quality) Then
   Call MsgBox("Fehler beim Erstellen von " & Ziel, vbExclamation)
  End If
  ConvertTojpg = FileLen(Ziel)
  Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ConvertTojpg/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ConvertTojpg&

Private Sub neutralisier()
  On Error GoTo fehler
'  PatInfosNichtNeu = -1
  Me.PatName = vNS
'  PatInfosNichtNeu = 0
  Me.KörperTeil = vNS
  Me.BildDatum.Value = Now
  Me.Beschreibung = vNS
  Me.Wagner = "-"
'  Me.Wagner.ListIndex = -1
'  Me.WagnerL = "-"
'  Me.Wagner.ListIndex = -1
  Me.Armstrong = vNS
  Me.NeuerName = vNS
  Me.PatName.BackColor = -2147483643
  Me.KörperTeil.BackColor = -2147483643
  Me.Beschreibung.BackColor = -2147483643
  Me.Wagner.BackColor = -2147483643
'  Me.WagnerL.BackColor = -2147483643
  Me.Armstrong.BackColor = -2147483643
 Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in neutralisier/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' neutralisier

Private Sub FDt_Ausgabe(Str$)
 Ausgabe = Ausgabe + Str
End Sub ' FDt_Ausgabe

Private Sub FDt_DBInit()
 Call DBInit
End Sub ' FDt_DBInit

Private Sub Form_Load()
 Set FDt = New FDatei
 Set dbv = New DBVerb
 Call do_Form_Load(Me)
 Me.Version = "Ver." & App.Major & "." & App.Minor & "." & App.Revision
End Sub ' Form_Load()

Private Sub Abbruch_Click()
 Dim erg&
 On Error GoTo fehler
 If FDC Is Nothing Then
 Else
'  erg = MsgBox("Der Abbruch-Knopf wurde betätigt. Wollen Sie vorher die Datensätze sichern?", vbYesNoCancel)
'  SELECT Case erg
'   Case vbYes: 'Call frmFDatei.aktSpeichern(Me)
'    Stop
'    Call FDC.Abbrechen
'   Case vbNo
'    Call FDC.Abbrechen
'   Case vbCancel: Exit Sub
'  End Select
' End If
  Call FDC.Abbrechen
 End If
 Unload Me
 End
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Abbruch_Click/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Abbruch_Click

Function max0(s1&)
 max0 = MAX(s1, 0)
End Function ' max0

Function MAX(s1&, s2&)
 If s1 > s2 Then
  MAX = s1
 Else
  MAX = s2
 End If
End Function ' MAX

Sub Form_Resize()
 On Error GoTo fehler
 Image1.Top = Ausgabe.Top + Ausgabe.Height + 20
 Image1.Height = max0(Me.Height - Image1.Top - 20)
 On Error Resume Next
 If GrößenVerhältnis = 0 Then GrößenVerhältnis = 1.3
 If Err.Number = 6 Then GrößenVerhältnis = 1.3
 On Error GoTo fehler
 Image1.Width = Image1.Height * GrößenVerhältnis '1.3
 On Error Resume Next
 Quelle.Width = max0(Me.Width * 0.5 - Quelle.Left) - 1975
 SteuerDB.Width = max0(Me.Width * 0.5 - SteuerDB.Left) - 1975
 On Error GoTo fehler
 BearbeiteteNochmalVerschieben.Left = CnStr.Left + CnStr.Width + 50
 obAcc.Left = DateiBreite.Left + DateiBreite.Width + 300
 obMySQL.Left = DateiBreite.Left + DateiBreite.Width + 300
' obMyQuelle1.Left = obAcc.Left
 SteuerDBBez.Left = Quelle.Left + Quelle.Width + 50
 SteuerDB.Left = SteuerDBBez.Left + SteuerDBBez.Width + 50
 Me.FtCn.Left = SteuerDB.Left + SteuerDB.Width + 10
 ArchPatBez.Left = FtCn.Left + FtCn.Width + 50
 ArchPat.Left = ArchPatBez.Left + ArchPatBez.Width + 50
 ArchPat.Width = max0(Me.Width - 200 - ArchPat.Left)
 ArchivBez.Left = ArchPatBez.Left
 Archiv.Left = ArchPat.Left
 Archiv.Width = ArchPat.Width
 NamDBBez.Left = ArchPatBez.Left
 Me.Version.Left = Me.NamDBBez.Left - 100 - Me.Version.Width
 Me.DateiPfad.Width = Me.Version.Left - Me.DateiPfad.Left - 100
 NamDB.Left = ArchPat.Left
 NamDB.Width = ArchPat.Width
 KomprBez.Left = ArchPatBez.Left
 Kompr.Left = ArchPat.Left
 Kompr.Width = ArchPat.Width
 Ausgabe.Width = max0(Me.Width - Ausgabe.Left - 200)
 PatName.Left = Image1.Left + Image1.Width + 50
 PatName.Width = max0(Me.Width - PatName.Left - 10)
 wieVoriges.Left = PatName.Left + PatName.Width - wieVoriges.Width - Me.wieNächstes.Width - 10
 Me.wieNächstes.Left = wieVoriges.Left + wieVoriges.Width + 10
 Me.nädP.Left = PatName.Left + PatName.Width - Me.nädP.Width
 Me.wieLetztesdPat.Left = PatName.Left + PatName.Width - Me.nädP.Width - Me.wieLetztesdPat.Width - 10
 Me.nur2.Left = PatName.Left + PatName.Width - Me.nädP.Width - Me.wieLetztesdPat.Width - Me.nur2.Width - 40
 SchrittweiteBez.Left = PatName.Left
 Schrittweite.Left = SchrittweiteBez.Left + SchrittweiteBez.Width + 20
 Position.Left = Schrittweite.Left + Schrittweite.Width + 70
 DSZahl.Left = Position.Left + Position.Width + 20
 ldP.Left = DSZahl.Left + DSZahl.Width + 40
 ndP.Left = ldP.Left + ldP.Width
 PatNamBez.Left = PatName.Left
 BildDatum.Left = PatName.Left
 BildDatumBez.Left = PatName.Left
 KörperTeil.Left = PatName.Left
 KörperTeilBez.Left = PatName.Left
 ZumAnfang.Left = PatName.Left
 RückwärtsCmd.Left = ZumAnfang.Left + ZumAnfang.Width + 20
 VorwärtsCmd.Left = RückwärtsCmd.Left + RückwärtsCmd.Width + 20
 ZumEnde.Left = VorwärtsCmd.Left + VorwärtsCmd.Width + 20
 ersterOffenerCmd.Left = ZumEnde.Left + ZumEnde.Width + 20
 Lad.Left = ersterOffenerCmd.Left + ersterOffenerCmd.Width + 20
 TonCmd.Left = Lad.Left + Lad.Width + 20 'DSZahl.Left + DSZahl.Width + 20
 keinTon.Left = TonCmd.Left - 620
 DTyp.Left = PatName.Left
 obPat.Left = DTyp.Left + DTyp.Width + 20
 WagnerBez.Left = PatName.Left
 Wagner.Left = WagnerBez.Left + WagnerBez.Width + 20
' WagnerL.Left = Wagner.Left + Wagner.Width + 20
 ArmstrongBez.Left = PatName.Left
 Armstrong.Left = ArmstrongBez.Left + ArmstrongBez.Width + 20
 BeschreibungBez.Left = PatName.Left
 Beschreibung.Left = PatName.Left
 Beschreibung.Width = max0(Me.Width - Beschreibung.Left - 10)
 LinksCmd.Left = PatName.Left
 rechtsCmd.Left = LinksCmd.Left + LinksCmd.Width + 100
 Grad.Left = PatName.Left
 FreiDrehen.Left = Grad.Left + Grad.Width + 100
 LinksNeu.Left = FreiDrehen.Left + FreiDrehen.Width + 100
 RechtsNeu.Left = LinksNeu.Left + LinksNeu.Width + 100
 
 HScroll1.Left = rechtsCmd.Left + rechtsCmd.Width + 100
 HScroll2.Left = rechtsCmd.Left + rechtsCmd.Width + 100
 HScroll3.Left = rechtsCmd.Left + rechtsCmd.Width + 100
 Me.FarbenZurück.Left = PatName.Left
 Me.HellLab.Left = Me.FarbenZurück.Left + Me.FarbenZurück.Width + 50
 Me.HelZ.Left = Me.HellLab.Left + Me.HellLab.Width + 100
 Me.KontrZ.Left = Me.HelZ.Left + Me.HelZ.Width + 20
 Me.GamZ.Left = Me.KontrZ.Left + Me.KontrZ.Width + 20
 DateiZeit.Left = BildDatum.Left + BildDatum.Width + 10
 AlleÜbertragen.Left = PatName.Left
 rueckgaengig.Left = AlleÜbertragen.Left + AlleÜbertragen.Width + 10
 stehenLassen.Left = AlleÜbertragen.Left + AlleÜbertragen.Width + 20
 KompressionsgradLbl.Left = stehenLassen.Left + stehenLassen.Width + 20
 Kompressionsgrad.Left = KompressionsgradLbl.Left + KompressionsgradLbl.Width + 20
 NeuerName.Left = PatName.Left
 NeuerName.Width = max0(Me.Width - NeuerName.Left - 10)
 FußstatusBez.Left = PatName.Left
 Fußstatus.Left = PatName.Left
 Dopplerlabel.Left = PatName.Left
 Doppler.Left = PatName.Left
 Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Form_Resize/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Form_Resize

Function FindDatei(Pfad$) As File
 Dim i&
 On Error GoTo f0
 For i = 0 To FDC.Count
  On Error GoTo fehler
  If FDC(i).Fil.Path = Pfad Then
   Set FindDatei = FDC(i)
   Exit For
  End If
 Next i
 Exit Function
f0:
If Err.Number = 9 Then Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in FindDatei/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' FindDatei

' kommt von in do_Start und findesatz
Function DBInit()
'  If SDB Is Nothing Then Set SDB = OpenDatabase(Me.SteuerDB)
'  If rSteu Is Nothing Then
'   Set rSteu = SDB.OpenRecordset("jpg", dbOpenTable)
'   rSteu.Index = "Name"
'  End If
End Function ' dbInit

Private Sub Abbruch_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' Abbruch_KeyDown

Private Sub AlleÜbertragen_Click()
 Call doAlleÜbertragen(Me)
' Call FDC.Item(FDC.indDat).DateiLad(Me, mitTon:=False)
End Sub ' AlleÜbertragen

Private Sub AlleÜbertragen_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' AlleÜbertragen_KeyDown

Private Sub Rueckgaengig_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' Rueckgaengig_KeyDown

'Private Sub AlteArchive_KeyDown(KeyCode As Integer, Shift As Integer)
' Call key(KeyCode, Shift, me)
'End Sub

Private Sub Archiv_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub Ausgabe_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub DateiBreite_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub DateiHöhe_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub DateiPfad_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub DateiZeit_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub Doppler_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub DSZahl_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub ersterOffenerCmd_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Call do_Form_Unload(Me)
End Sub

Private Sub FreiDrehen_Click()
 Call RotatePicDI(Picture2, Picture2, Grad)
End Sub

Private Sub FtCn_Click()
 Call dbv.Auswahl(vNS, "jpg", "fotosinp")
End Sub

Private Sub LinksNeu_Click()
 Call RotatePicDI(Picture2, Picture2, 270)
End Sub

Private Sub nur2_Click()
' nur zwei Wörter von übernommenem Text dalassen
 Dim pos%
 pos = InStr(Me.PatName, " ")
 If pos > 0 Then
  pos = InStr(pos + 1, Me.PatName, " ")
  Me.PatName = Left(Me.PatName, pos)
 End If
 Me.PatName.SetFocus
 Me.PatName.SelStart = Len(Me.PatName)
End Sub ' nur2_Click()

Private Sub obMySQL_Click()
' Call dbv.cnVorb(vNS, "anamnesebogen", "Quelle")
 If Not imAufbau Then
  Call ConstrFestleg(Me, 2)
 End If
End Sub ' obMySQL_Click()

Private Sub RechtsNeu_Click()
 Call RotatePicDI(Picture2, Picture2, 90)
End Sub

Private Sub Fußstatus_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub GamZ_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub HelZ_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub Kompr_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub KontrZ_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub Lad_Click()
 Call doPatNameChange(Me)
End Sub

Private Sub Lad_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub LinksCmd_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub nur2_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub LinksNeu_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub RechtsNeu_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub FreiDrehen_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub Grad_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub NamDB_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub NeuerName_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub FarbenZurück_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub BildDatum_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub obAcc_Click()
 If Not imAufbau Then
  Call ConstrFestleg(Me, 1)
  Call AuswName(Me)
'  Me.PatName.Refresh
'  Me.Refresh
 End If
End Sub ' obAcc_Click

Private Sub obMyQuelle_Click()
 If Not imAufbau Then
  Call ConstrFestleg(Me, 2)
  Call AuswName(Me)
'  Me.PatName.Refresh
'  Me.Refresh
 End If
End Sub

Private Sub obMyQuelle1_Click()
 If Not imAufbau Then
  Call ConstrFestleg(Me, 3)
  Call AuswName(Me)
'  Me.PatName.Refresh
'  Me.Refresh
 End If
End Sub

Private Sub PatName_Validate(Cancel As Boolean)
 Call doPatNameChange(Me)
End Sub

Private Sub Position_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub rechtsCmd_Click()
 Call Rotier(jpgTransformrotate90, Me)
End Sub

Private Sub LinksCmd_Click()
 Call Rotier(jpgTransformrotate270, Me)
End Sub

Private Sub rechtsCmd_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub HelZ_GotFocus()
 Me.HelZ.SelStart = 0
 Me.HelZ.SelLength = Len(Me.HelZ)
End Sub ' HelZ_GotFocus()

Private Sub Schrittweite_GotFocus()
 Me.Schrittweite.SelStart = 0
 Me.Schrittweite.SelLength = Len(Me.Schrittweite)
End Sub ' Schrittweite_GotFocus()

Private Sub Schrittweite_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' Schrittweite_KeyDown(KeyCode As Integer, Shift As Integer)

'Private Sub obAlteArchive_KeyDown(KeyCode As Integer, Shift As Integer)
' Call key(KeyCode, Shift, me)
'End Sub

'Private Sub obEingel_KeyDown(KeyCode As Integer, Shift As Integer)
' Call key(KeyCode, Shift, me)
'End Sub

'Private Sub obQuelle_KeyDown(KeyCode As Integer, Shift As Integer)
' Call key(KeyCode, Shift, me)
'End Sub


Private Sub SteuerDB_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' SteuerDB_KeyDown

Private Sub Quelle_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' Quelle_KeyDown

Private Sub Start_Click()
 Call do_Start(Me)
End Sub ' Start_Click

Private Sub Lw_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' Lw_KeyDown

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

'Private Sub Eingel_KeyDown(KeyCode As Integer, Shift As Integer)
' Call key(KeyCode, Shift, me)
'End Sub


Private Sub Start_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' Start_KeyDown

Private Sub ArchPat_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' ArchPat_KeyDown

Private Sub TonCmd_Click()
 Call Sound(FDt.wavPfad)
End Sub ' TonCmd_Click

Private Sub TonCmd_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' TonCmd_KeyDown

Private Sub obpat_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' obpat_KeyDown

Private Sub ldP_Click()
 Call doFindeLetztendPat(Me)
End Sub ' ldP_Click

Private Sub ndP_Click()
 Call doFindeNächstendPat(Me)
End Sub ' ndP_Click

Private Sub VorwärtsCmd_Click()
 If FDC.indDat < Me.FDC.Count Then
  Call doVorwärtsCmd(Me)
 Else
  Beep
 End If
End Sub ' VorwärtsCmd

Private Sub ersterOffenerCmd_Click()
 Call doersterOffenerCmd(Me)
End Sub ' ersterOffenerCmd_Click

Private Sub Wagner_Change()
' Stop
End Sub ' Wagner_Change

Private Sub Wagner_KeyPress(KeyAscii As Integer)
 Call Key(KeyAscii, 0, Me)
End Sub ' Wagner_KeyPress

Private Sub Wagner_Validate(Cancel As Boolean)
' Stop
End Sub ' Wagner_Validate

'Private Sub WagnerL_Validate(Cancel As Boolean)
' Stop
'End Sub
Private Sub wieNächstes_Click()
 Call doWieNächstes(Me)
 Call doPatNameChange(Me)
End Sub ' wieNächstes_Click

Private Sub wieLetztesdPat_Click()
 Call doWieLetztesdPat(Me)
 Call doPatNameChange(Me)
End Sub ' wieLetztesdPat_Click

Private Sub nädP_Click()
 Call doWieLetztesdPat(Me, True)
 Call doPatNameChange(Me)
End Sub ' nädP_Click

Private Sub wieVoriges_Click()
 Call doWieVoriges(Me)
 Call doPatNameChange(Me)
 Me.PatName.SetFocus
 On Error Resume Next
 Me.PatName.SelStart = Len(Me.PatName)
End Sub ' wieVoriges_Click

Private Sub wieletztesdpat_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' wieletztesdpat_KeyDown

Private Sub nädP_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' nädP_KeyDown

Private Sub wieNächstes_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' wieNächstes_KeyDown

Private Sub wieVoriges_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' wieVoriges_KeyDown

Private Sub ZumAnfang_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' ZumAnfang_KeyDown

Private Sub ZumEnde_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' ZumEnde_KeyDown

Private Sub rückwärtsCmd_Click()
 If FDC.indDat > 0 Then
  Call doRückwärtsCmd(Me)
 Else
  Beep
 End If
End Sub ' rückwärtsCmd_Click

Private Sub zumanfang_click()
 Me.FDC.indDat = 1
' Call FDC(FDC.indDat).DateiLad(Me, mitTon:=False)
End Sub ' zumanfang_click

Private Sub zumende_click()
 FDC.indDat = Me.FDC.Count
' Call FDC(FDC.indDat).DateiLad(Me, mitTon:=False)
End Sub

Private Sub Wagner_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub Beschreibung_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub Armstrong_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub ldP_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub ndP_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub rückwärtsCmd_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub vorwärtsCmd_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub patname_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub Körperteil_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub obMyQuelle_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub obMyQuelle1_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub obAcc_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub HScroll1_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub HScroll2_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub HScroll3_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub

Private Sub HScroll1_Change()
 Call HellKontr
 Me.HelZ = Me.HScroll1
End Sub ' HScroll1_Change

Private Sub HScroll2_Change()
 Call HellKontr
 Me.KontrZ = Me.HScroll2
End Sub ' HScroll2_Change

Private Sub HScroll3_Change()
 Call HellKontr
 Me.GamZ = Me.HScroll3
End Sub ' HScroll3_Change

Private Sub HelZ_validate(Cancel As Boolean)
 If Me.HScroll1 <> Me.HelZ Then
  Me.HScroll1 = Me.HelZ
 End If
End Sub ' HelZ_validate

Private Sub kontrz_validate(Cancel As Boolean)
 If Me.HScroll2 <> Me.KontrZ Then
  Me.HScroll2 = Me.KontrZ
 End If
End Sub ' kontrz_validate

Private Sub gamz_validate(Cancel As Boolean)
 If Me.HScroll3 <> Me.GamZ Then
  Me.HScroll3 = Me.GamZ
 End If
End Sub ' gamz_validate

Public Sub HellKontr()
  Dim x%, Temp!, TempAlt!
   If Not BildNichtNeu Then
    Screen.MousePointer = vbHourglass
    neuContrast = Exp(HScroll2.Value / 30) / 20 - 0.05
    neuBrightness = Exp(HScroll1.Value / 50) / 5 - 0.2
    neuGamma = (250 / HScroll3.Value) ^ 4
    If neuContrast <> Contrast Or neuBrightness <> Brightness Or neuGamma <> Gamma Then
     Brightness = neuBrightness
     Contrast = neuContrast
     Gamma = neuGamma
     For x = 0 To 255
      TempAlt = ((x * Brightness - 127) * Contrast) + 127
      If TempAlt > 255 Then TempAlt = 255
      If TempAlt < 0 Then TempAlt = 0
      Temp = ((TempAlt / 255) ^ Gamma) * 255
      If Temp > 255 Then Temp = 255
      If Temp < 0 Then Temp = 0
      TableBright(x) = Temp
     Next x
     Call MakeBitmap
     Image1.Refresh
    End If
    Screen.MousePointer = 0
 End If
 Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in HellKontr/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' HellKontr

Public Sub MakeBitmap()
  Dim Pic() As Byte, PicBuff() As Byte
  Dim SA As SAFEARRAY2D, bmp As BITMAP
  Dim SABuff As SAFEARRAY2D, BmpBuff As BITMAP
  Dim x%, Y%

    Call GetObjectgdi32(Image1.Picture, Len(bmp), bmp)
    Call GetObjectgdi32(Picture2.Picture, Len(BmpBuff), BmpBuff)
  
    If bmp.bmBitsPixel <> 24 Then
      MsgBox "Es werden nur 24-Bit Bitmaps unterstützt!"
      Exit Sub
    End If
    
    With SA
      .cbElements = 1
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = bmp.BmHeight
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = bmp.bmWidthBytes
      .pvData = bmp.bmBits
    End With
    
    Call CopyMemory(ByVal VarPtrArray(Pic), VarPtr(SA), 4)
        
    With SABuff
      .cbElements = 1
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = BmpBuff.BmHeight
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = BmpBuff.bmWidthBytes
      .pvData = BmpBuff.bmBits
    End With
    
    Call CopyMemory(ByVal VarPtrArray(PicBuff), VarPtr(SABuff), 4)
'    On Error Resume Next
    For x = 0 To UBound(Pic, 1)
      For Y = 0 To UBound(Pic, 2)
        Pic(x, Y) = TableBright(PicBuff(x, Y))
      Next Y
    Next x
'    On Error GoTo fehler
    Call CopyMemory(ByVal VarPtrArray(Pic), 0&, 4)
    Call CopyMemory(ByVal VarPtrArray(PicBuff), 0&, 4)
 Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in MakeBitmap/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' MakeBitmap

