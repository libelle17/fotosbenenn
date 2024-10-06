Attribute VB_Name = "AFastRotate"
Option Explicit
Public SrcPix() As Byte, DesPix() As Byte
'A Fast Rotate

'**************************************
'Windows API/Global Declarations for :A
'     Fast Rotate
'**************************************


Type BITMAPINFOHEADER '40 bytes
    BmSize As Long
    BmWidth As Long
    BmHeight As Long
    BmPlanes As Integer
    BmBitCount As Integer
    BmCompression As Long
    BmSizeImage As Long
    BmXPelsPerMeter As Long
    BmYPelsPerMeter As Long
    BmClrUsed As Long
    BmClrImportant As Long
    End Type


Type BITMAPINFO
    BmHeader As BITMAPINFOHEADER
    End Type
    'VB 16

'Declare Sub GetDIBits Lib "GDI" (ByVal hDC%, ByVal hBitmap%, ByVal nStartScan%, ByVal nNumScans%, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage%)


'Declare Sub SetDIBits Lib "GDI" (ByVal hDC%, ByVal hBitmap%, ByVal nStartScan%, ByVal nNumScans%, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage%)
    'VB 32


Declare Sub GetDIBits Lib "GDI32" (ByVal hDC&, ByVal hBitmap&, ByVal nStartScan&, ByVal nNumScans&, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage&)


Declare Sub SetDIBits Lib "GDI32" (ByVal hDC&, ByVal hBitmap&, ByVal nStartScan&, ByVal nNumScans&, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage&)


'**************************************
' Name: A Fast Rotate
' Description:Rotates a Picture USING GetDIBits/SetDIBits Very Fast
' By: Show
'
'
' Inputs:None
'
' Returns:None
'
'Assumes:Add 2 PictureBoxs Load A Picture In 1
'AutoSize True
'Both (Pixels)
'
'Side Effects:None
'This code is copyrighted and has limited warranties.
'Please see http://www.Planet-Source-Code.com/xq/ASP/txtCodeId.51353/lngWId.1/qx/vb/scripts/ShowCode.htm
'for details.
'**************************************


Sub RotatePicDI(SrcPic As PictureBox, DestPic As PictureBox, A As Double)
    Dim SrcInfo As BITMAPINFO, DesInfo As BITMAPINFO
    Dim x&, Y&, CA As Double, SA As Double, nX&, nY&
    Dim sW&, sH&, sW2&, sH2&, dw&, dH&, dW2&, dH2&
    'RotatePicDI Picture1, Picture2, 45
    Const Pi = 0.017453292519943
    CA = Cos(A * Pi * -1): SA = Sin(A * Pi * -1)
    sW = SrcPic.ScaleWidth
    sH = SrcPic.ScaleHeight
    dw = DestPic.ScaleWidth
    dH = DestPic.ScaleHeight
    sW2 = sW / 2: sH2 = sH / 2
    dW2 = dw / 2: dH2 = dH / 2
    SrcInfo.BmHeader.BmSize = 40 'Always
    SrcInfo.BmHeader.BmWidth = sW 'Width
    SrcInfo.BmHeader.BmHeight = -sH 'If You Want to Start Top-Botttom Put -Height
    SrcInfo.BmHeader.BmPlanes = 1 'Always
    SrcInfo.BmHeader.BmBitCount = 32 ' Can Be 16, 24, 32; 32 Is Best
    SrcInfo.BmHeader.BmSizeImage = 3 * sW * sH
    'If You Change The BitCount To 16 Or 24
    'You Have To Change The SrcPix And DesPix Values
    'Example: ReDim SrcPix("0,1,2,3,4" , sW - 1, sH - 1) As Long
    'I Think For VB32 Users If You Have BitCount 32
    'You Have To Change SrcPix And DesPix Values To 3,W,H
    '(ReDim SrcPix(3, sW - 1, sH - 1) As Byte)
    'Or (ReDim SrcPix(0 To 2, sW - 1, sH - 1) As Byte)
    'This Should Get You The Red,Green,Blue Values
    '2=Red,1=Green,0=Blue | 3=Red,2=Green,1= Blue
    LSet DesInfo = SrcInfo 'Copy SrcInfo to DesInfo
    DesInfo.BmHeader.BmWidth = dw 'Width
    DesInfo.BmHeader.BmHeight = -dH 'If You Want to
    DesInfo.BmHeader.BmSizeImage = 3 * dw * dH
    'Start Top-Botttom Put -Height
    ReDim DesPix(0, dw - 1, dH - 1)
    ReDim SrcPix(0, sW - 1, sH - 1)
    'Dont work try this
    'ReDim SrcPix(0, sW - 1, sH - 1) As Byte
    'ReDim DesPix(0, dW - 1, dH - 1) As Byte
    '
    'Or this
    'ReDim SrcPix(0 To 2, sW - 1, sH - 1) As Byte
    'ReDim DesPix(0 To 2, dW - 1, dH - 1) As Byte
    'Also You Might Have To Change
    'Pic.Image To Pic.Image.Handle
    'Call GetDIBits(SrcPic.hDC, SrcPic.Image
    '     .Handle, 0&, sH, SrcPix(0, 0, 0), SrcInf
    '     o, 0&)
    Call GetDIBits(SrcPic.hDC, SrcPic.image, 0&, sH, SrcPix(0, 0, 0), SrcInfo, 0&)


    For Y = 0 To dH - 1


        For x = 0 To dw - 1
            nX = CA * (x - dW2) - SA * (Y - dH2) + sW2
            nY = SA * (x - dW2) + CA * (Y - dH2) + sH2


            If nX > -1 And nY > -1 And nX < sW And nY < sH Then
                DesPix(0, x, Y) = SrcPix(0, nX, nY)
                'VB32 Might Have To Use This
                'DesPix(1, X, Y) = SrcPix(1, nX, nY)
                'DesPix(2, X, Y) = SrcPix(2, nX, nY)
            End If
        Next
    Next
    'Call SetDIBits(DestPic.hDC, DestPic.Ima
    '     ge.Handle, 0&, dH, DesPix(0, 0, 0), DesI
    '     nfo, 0&)
    Call SetDIBits(DestPic.hDC, DestPic.image, 0&, dH, DesPix(0, 0, 0), DesInfo, 0&)
    DestPic.Picture = DestPic.image
End Sub
        

