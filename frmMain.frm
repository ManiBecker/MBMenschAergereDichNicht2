VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "MBMenschAergereDichNicht2"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6210
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MBMenschAergereDichNicht2"
   ScaleHeight     =   392
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   414
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Shape shpSpieler 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      Height          =   375
      Index           =   0
      Left            =   4200
      Shape           =   3  'Kreis
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape shpStart 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   0
      Left            =   3300
      Shape           =   3  'Kreis
      Top             =   780
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H000000FF&
      Height          =   495
      Index           =   0
      Left            =   2460
      Shape           =   3  'Kreis
      Top             =   780
      Width           =   495
   End
   Begin VB.Image ImageBackground 
      Appearance      =   0  '2D
      Height          =   1995
      Left            =   0
      Picture         =   "frmMain.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1995
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////
'//
'// Projekt : MBMenschAergereDichNicht2
'// Sprache : Basic
'// Compiler: MS Visual Basic 6.0
'// Autor   : Manfred Becker
'// E-Mail  : mani.becker@web.de
'// Url     : https://github.com/ManiBecker/MBMenschAergereDichNicht2
'// Modul   : frmMain
'// Version : 1.00
'// Datum   : 24.05.2020
'//
'///////////////////////////////////////////////////////////////////////////

Option Explicit

Dim BmpWidth!, BmpHeight!, BmpDiv!
Dim ImgWidth!, ImgHeight!, ImgTop!, ImgLeft!
Dim FrmWidth!, FrmHeight!
Dim dx!, dy!, bw!
Dim mdX!, mdY!
Dim mdLeft!
Dim dXStart!, dYStart!
Dim fx!
Dim mfx(55) As Integer
Dim mfy(55) As Integer
Dim msx(15) As Integer
Dim msy(15) As Integer
Dim mpx(15) As Integer
Dim mpy(15) As Integer

Dim colorPlayerName(4) As String
Dim colorPlayer(4) As Long
Dim colorStartField(4) As Long
Dim colorTargetField(4) As Long
Const vbLightGreen As Long = &HC0FFC0
Const vbLightRed As Long = &HC0C0FF
Const vbLightBlack As Long = &HC0C0C0
Const vbLightYellow As Long = &HC0FFFF

Dim bBitmapFound As Boolean
Dim bResize As Boolean
Dim bMove As Boolean
Dim bClick As Boolean
Public Enum enmGameStatus
    Initial
    Running
    Stopped
End Enum
Dim Gamestatus As enmGameStatus

Private Sub RePaint()
    Dim n%
    
    If Not bBitmapFound Then
        DrawMode = vbCopyPen
    
        Me.Cls
        
        
        DrawMode = vbXorPen
        
    End If
End Sub


Private Sub Form_Load()

On Error GoTo ErrorHandler

    Gamestatus = enmGameStatus.Initial

    'Maßeinheit auf Pixel einstellen
    Me.ScaleMode = vbPixels
    
    'zuerst Bitmap laden...
    bBitmapFound = True
    'ImageBackground.Picture = Nothing 'LoadPicture("img_background.jpg")
    bBitmapFound = False
    
    If bBitmapFound Then
        'nun Orginalgrösse der Bitmap ermitteln...
        BmpWidth = ImageBackground.Width
        BmpHeight = ImageBackground.Height
        BmpDiv = BmpHeight / BmpWidth
        AutoRedraw = False
        DrawMode = vbXorPen
    Else
        BmpWidth = 133
        BmpHeight = 133
        BmpDiv = BmpHeight / BmpWidth
        AutoRedraw = False
        DrawMode = vbXorPen
    End If
    
    'jetzt Picture-Abmessung anpassen, der Rest erledigt Form_Resize()...
    ImageBackground.Stretch = True
    
    FrmWidth = 0
    FrmHeight = 0
    
    'Matrix Koordinaten der Spielfeldpunkte 0 bis 39
    mfx(0) = 6: mfy(0) = 0
    mfx(1) = 6: mfy(1) = 1
    mfx(2) = 6: mfy(2) = 2
    mfx(3) = 6: mfy(3) = 3
    mfx(4) = 6: mfy(4) = 4
    mfx(5) = 7: mfy(5) = 4
    mfx(6) = 8: mfy(6) = 4
    mfx(7) = 9: mfy(7) = 4
    mfx(8) = 10: mfy(8) = 4
    mfx(9) = 10: mfy(9) = 5
    mfx(10) = 10: mfy(10) = 6
    mfx(11) = 9: mfy(11) = 6
    mfx(12) = 8: mfy(12) = 6
    mfx(13) = 7: mfy(13) = 6
    mfx(14) = 6: mfy(14) = 6
    mfx(15) = 6: mfy(15) = 7
    mfx(16) = 6: mfy(16) = 8
    mfx(17) = 6: mfy(17) = 9
    mfx(18) = 6: mfy(18) = 10
    mfx(19) = 5: mfy(19) = 10
    mfx(20) = 4: mfy(20) = 10
    mfx(21) = 4: mfy(21) = 9
    mfx(22) = 4: mfy(22) = 8
    mfx(23) = 4: mfy(23) = 7
    mfx(24) = 4: mfy(24) = 6
    mfx(25) = 3: mfy(25) = 6
    mfx(26) = 2: mfy(26) = 6
    mfx(27) = 1: mfy(27) = 6
    mfx(28) = 0: mfy(28) = 6
    mfx(29) = 0: mfy(29) = 5
    mfx(30) = 0: mfy(30) = 4
    mfx(31) = 1: mfy(31) = 4
    mfx(32) = 2: mfy(32) = 4
    mfx(33) = 3: mfy(33) = 4
    mfx(34) = 4: mfy(34) = 4
    mfx(35) = 4: mfy(35) = 3
    mfx(36) = 4: mfy(36) = 2
    mfx(37) = 4: mfy(37) = 1
    mfx(38) = 4: mfy(38) = 0
    mfx(39) = 5: mfy(39) = 0
    'Matrix Koordinaten der grünen Zielpunkte
    mfx(40) = 5: mfy(40) = 1
    mfx(41) = 5: mfy(41) = 2
    mfx(42) = 5: mfy(42) = 3
    mfx(43) = 5: mfy(43) = 4
    'Matrix Koordinaten der roten Zielpunkte
    mfx(44) = 9: mfy(44) = 5
    mfx(45) = 8: mfy(45) = 5
    mfx(46) = 7: mfy(46) = 5
    mfx(47) = 6: mfy(47) = 5
    'Matrix Koordinaten der schwarzen Zielpunkte
    mfx(48) = 5: mfy(48) = 9
    mfx(49) = 5: mfy(49) = 8
    mfx(50) = 5: mfy(50) = 7
    mfx(51) = 5: mfy(51) = 6
    'Matrix Koordinaten der gelben Zielpunkte
    mfx(52) = 1: mfy(52) = 5
    mfx(53) = 2: mfy(53) = 5
    mfx(54) = 3: mfy(54) = 5
    mfx(55) = 4: mfy(55) = 5
    'Matrix Koordinaten der grünen Startpunkte
    msx(0) = 9: msy(0) = 0
    msx(1) = 10: msy(1) = 0
    msx(2) = 9: msy(2) = 1
    msx(3) = 10: msy(3) = 1
    'Matrix Koordinaten der roten Startpunkte
    msx(4) = 9: msy(4) = 9
    msx(5) = 10: msy(5) = 9
    msx(6) = 9: msy(6) = 10
    msx(7) = 10: msy(7) = 10
    'Matrix Koordinaten der schwarzen Startpunkte
    msx(8) = 0: msy(8) = 9
    msx(9) = 1: msy(9) = 9
    msx(10) = 0: msy(10) = 10
    msx(11) = 1: msy(11) = 10
    'Matrix Koordinaten der gelben Startpunkte
    msx(12) = 0: msy(12) = 0
    msx(13) = 1: msy(13) = 0
    msx(14) = 0: msy(14) = 1
    msx(15) = 1: msy(15) = 1
    'Matrix Koordinaten der grünen Spieler
    mpx(0) = 9: mpy(0) = 0
    mpx(1) = 10: mpy(1) = 0
    mpx(2) = 9: mpy(2) = 1
    mpx(3) = 10: mpy(3) = 1
    'Matrix Koordinaten der roten Spieler
    mpx(4) = 9: mpy(4) = 9
    mpx(5) = 10: mpy(5) = 9
    mpx(6) = 9: mpy(6) = 10
    mpx(7) = 10: mpy(7) = 10
    'Matrix Koordinaten der schwarzen Spieler
    mpx(8) = 0: mpy(8) = 9
    mpx(9) = 1: mpy(9) = 9
    mpx(10) = 0: mpy(10) = 10
    mpx(11) = 1: mpy(11) = 10
    'Matrix Koordinaten der gelben Spieler
    mpx(12) = 0: mpy(12) = 0
    mpx(13) = 1: mpy(13) = 0
    mpx(14) = 0: mpy(14) = 1
    mpx(15) = 1: mpy(15) = 1
    
    
    colorPlayerName(0) = "Grün"
    colorPlayerName(1) = "Rot"
    colorPlayerName(2) = "Schwarz"
    colorPlayerName(3) = "Gelb"
    colorPlayer(0) = vbGreen
    colorPlayer(1) = vbRed
    colorPlayer(2) = vbBlack
    colorPlayer(3) = vbYellow
    colorStartField(0) = vbLightGreen
    colorStartField(1) = vbLightRed
    colorStartField(2) = vbLightBlack
    colorStartField(3) = vbLightYellow
    colorTargetField(0) = vbLightGreen
    colorTargetField(1) = vbLightRed
    colorTargetField(2) = vbLightBlack
    colorTargetField(3) = vbLightYellow

    
    Dim i, j As Integer
    
    'Spielfelder
    For i = 0 To 39
        If i > 0 Then Load shpFeld(i)
        With shpFeld(i)
            .Visible = True
            .BorderStyle = 1
            .FillStyle = 1
            .BorderColor = vbBlack
            .ZOrder 0
            If bBitmapFound Then
                .BackStyle = 0
            Else
                .BackStyle = 1
                If i = 0 Then
                    .BackColor = colorTargetField(0)
                ElseIf i = 10 Then
                    .BackColor = colorTargetField(1)
                ElseIf i = 20 Then
                    .BackColor = colorTargetField(2)
                ElseIf i = 30 Then
                    .BackColor = colorTargetField(3)
                Else
                    .BackColor = vbWhite
                End If
            End If
        End With
    Next i

    'Zielfelder
    For j = 0 To 3
    For i = i To i + 3
        Load shpFeld(i)
        With shpFeld(i)
            .Visible = True
            .BorderStyle = 1
            .FillStyle = 1
            .ZOrder 0
            If bBitmapFound Then
                .BackStyle = 0
            Else
                .BackStyle = 1
                .BackColor = colorTargetField(j)
            End If
        End With
    Next i
    Next j
    
    'Startfelder
    i = 0
    For j = 0 To 3
    For i = i To i + 3
        If i > 0 Then Load shpStart(i)
        With shpStart(i)
            .Visible = True
            .BorderStyle = 1
            .FillStyle = 1
            .ZOrder 0
            If bBitmapFound Then
                .BackStyle = 0
            Else
                .BackStyle = 1
                .BackColor = colorStartField(j)
            End If
        End With
    Next i
    Next j
    
    'Spieler
    i = 0
    For j = 0 To 3
    For i = i To i + 3
        If i > 0 Then Load shpSpieler(i)
        With shpSpieler(i)
            .Visible = True
            .BorderStyle = 1
            .FillStyle = 1
            .ZOrder 0
            .BackStyle = 1
            .BackColor = colorPlayer(j)
        End With
    Next i
    Next j
    
    Exit Sub
    
ErrorHandler:
    Dim ErrNumber, ErrDescription
    
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    
    Select Case ErrNumber
        Case 53
            MsgBox ErrDescription, , "Hinweis"
            bBitmapFound = False
            Resume Next
        Case Else
            MsgBox ErrDescription, , "Fehler " & ErrNumber

    End Select
    
End Sub

Private Sub Form_Resize()
    If WindowState = vbMinimized Then Exit Sub
    
    Dim i As Integer
    
    'minimale Größe festlegen
    If Me.Width < 2400 Then Me.Width = 2400
    If Me.Height < 2000 Then Me.Height = 2000
    
    'berechne Image-Grösse...
    If Me.Width < (Me.Height - 300) Then
        ImgWidth = Me.Width / Screen.TwipsPerPixelX - 8
        ImgHeight = ImgWidth * BmpDiv
        ImgTop = (Me.Height / Screen.TwipsPerPixelY - 32 - ImgHeight) / 2
        ImgLeft = 0
    Else
        ImgWidth = Me.Height / Screen.TwipsPerPixelY - 32
        ImgHeight = ImgWidth * BmpDiv
        ImgTop = 0
        ImgLeft = (Me.Width / Screen.TwipsPerPixelX - 8 - ImgWidth) / 2
    End If
    
    'jetzt Picture-Abmessung anpassen...
    ImageBackground.Top = ImgTop
    ImageBackground.Left = ImgLeft
    ImageBackground.Width = ImgWidth
    ImageBackground.Height = ImgHeight
    
    'Abmessungen der Form sichern...
    FrmWidth = Me.Width
    FrmHeight = Me.Height
    
    dx = ImgWidth / BmpWidth
    dy = ImgHeight / BmpHeight
    
    If dx > 2 Then
        bw = dx / 2
    Else
        bw = 1
    End If
    For i = 0 To 55
        If i < 40 Then
            shpFeld(i).Left = ImgLeft + dx * 8 + dx * mfx(i) * 10.9
            shpFeld(i).Top = ImgTop + dy * 8 + dy * mfy(i) * 10.9
            shpFeld(i).Width = dx * 8.5
            shpFeld(i).Height = dy * 8.5
            shpFeld(i).BorderWidth = bw
        Else
            shpFeld(i).Left = ImgLeft + dx * 8 + dx * mfx(i) * 10.9 + dx
            shpFeld(i).Top = ImgTop + dy * 8 + dy * mfy(i) * 10.9 + dy
            shpFeld(i).Width = dx * 6.5
            shpFeld(i).Height = dy * 6.5
            shpFeld(i).BorderWidth = bw
        End If
        If i < 16 Then
            shpStart(i).Left = ImgLeft + dx * 8 + dx * msx(i) * 10.9 + dx
            shpStart(i).Top = ImgTop + dy * 8 + dy * msy(i) * 10.9 + dy
            shpStart(i).Width = dx * 6.5
            shpStart(i).Height = dy * 6.5
            shpStart(i).BorderWidth = bw
            
            shpSpieler(i).Left = ImgLeft + dx * 8 + dx * msx(i) * 10.9 + dx
            shpSpieler(i).Top = ImgTop + dy * 8 + dy * msy(i) * 10.9 + dy
            shpSpieler(i).Width = dx * 6.5
            shpSpieler(i).Height = dy * 6.5
            shpSpieler(i).BorderWidth = bw
        End If
    Next i
    
    
    If Not bBitmapFound Then
        Call RePaint
    End If
End Sub

