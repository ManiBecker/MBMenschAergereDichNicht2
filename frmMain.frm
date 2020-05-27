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
Dim dx!, dy!
Dim mdX!, mdY!
Dim mdLeft!
Dim dXStart!, dYStart!
Dim fx!
Dim mx(41) As Integer
Dim my(41) As Integer

Dim bBitmapFound As Boolean
Dim bResize As Boolean
Dim bMove As Boolean
Dim bClick As Boolean

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

    'Maßeinheit auf Pixel einstellen
    Me.ScaleMode = vbPixels
    
    'zuerst Bitmap laden...
    bBitmapFound = True
    ImageBackground.Picture = LoadPicture("img_background.jpg")
    'bBitmapFound = False
    
    If bBitmapFound Then
        'nun Orginalgrösse der Bitmap ermitteln...
        BmpWidth = ImageBackground.Width
        BmpHeight = ImageBackground.Height
        BmpDiv = BmpHeight / BmpWidth
        AutoRedraw = False
        DrawMode = vbXorPen
    Else
        BmpWidth = 2000
        BmpHeight = 2000
        BmpDiv = BmpHeight / BmpWidth
        AutoRedraw = True
        DrawMode = vbXorPen
    End If
    
    'jetzt Picture-Abmessung anpassen, der Rest erledigt Form_Resize()...
    ImageBackground.Stretch = True
    
    FrmWidth = 0
    FrmHeight = 0
    
    mx(0) = 6: my(0) = 0
    mx(1) = 6: my(1) = 1
    mx(2) = 6: my(2) = 2
    mx(3) = 6: my(3) = 3
    mx(4) = 6: my(4) = 4
    mx(5) = 7: my(5) = 4
    mx(6) = 8: my(6) = 4
    mx(7) = 9: my(7) = 4
    mx(8) = 10: my(8) = 4
    mx(9) = 10: my(9) = 5
    mx(10) = 10: my(10) = 6
    mx(11) = 9: my(11) = 6
    mx(12) = 8: my(12) = 6
    mx(13) = 7: my(13) = 6
    mx(14) = 6: my(14) = 6
    mx(15) = 6: my(15) = 7
    mx(16) = 6: my(16) = 8
    mx(17) = 6: my(17) = 9
    mx(18) = 6: my(18) = 10
    mx(19) = 5: my(19) = 10
    mx(20) = 4: my(20) = 10
    mx(21) = 4: my(21) = 9
    mx(22) = 4: my(22) = 8
    mx(23) = 4: my(23) = 7
    mx(24) = 4: my(24) = 6
    mx(25) = 3: my(25) = 6
    mx(26) = 2: my(26) = 6
    mx(27) = 1: my(27) = 6
    mx(28) = 0: my(28) = 6
    mx(29) = 0: my(29) = 5
    mx(30) = 0: my(30) = 4
    mx(31) = 1: my(31) = 4
    mx(32) = 2: my(32) = 4
    mx(33) = 3: my(33) = 4
    mx(34) = 4: my(34) = 4
    mx(35) = 4: my(35) = 3
    mx(36) = 4: my(36) = 2
    mx(37) = 4: my(37) = 1
    mx(38) = 4: my(38) = 0
    mx(39) = 5: my(39) = 0
    
    Dim i As Integer
    
    For i = 0 To 39
        If i > 0 Then Load shpFeld(i)
        With shpFeld(i)
            .Visible = True
            .BackStyle = 0
            .BorderStyle = 1
            .FillStyle = 1
            .BorderColor = vbBlack
            .ZOrder 0
        End With
    Next i

    
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
    
    'nun noch die Höhe der Form anpassen...
'    Me.Height = Screen.TwipsPerPixelY * ImageBackground.Height + 400
    
    'Abmessungen der Form sichern...
    FrmWidth = Me.Width
    FrmHeight = Me.Height
    
    dx = ImgWidth / BmpWidth
    dy = ImgHeight / BmpHeight
    
    For i = 0 To shpFeld.Count - 1
        shpFeld(i).Left = ImgLeft + dx * 8 + dx * mx(i) * 10.9
        shpFeld(i).Top = ImgTop + dy * 8 + dy * my(i) * 10.9
        shpFeld(i).Width = dx * 8.5
        shpFeld(i).Height = dy * 8.5
        If dx > 2 Then
            shpFeld(i).BorderWidth = dx / 2
        Else
            shpFeld(i).BorderWidth = 1
        End If
    Next i
    Me.Caption = dx & "/" & dy
    
'    Label1.Left = 150 * dx * fx
'    Label1.Top = 130 * dy * fx
'    Label1.Width = 1250 * dx * fx
'    Label1.Height = 70 * dy * fx
'    Label1.Font.Size = Int(Label1.Width / 32) * 32 / 29
'
'    Label2.Left = Label1.Left
'    Label2.Top = Label1.Top + Label1.Height
'    Label2.Width = Label1.Width
'    Label2.Height = Label1.Height
'    Label2.Font.Size = Label1.Font.Size
'
'    Me.Font.Size = Label1.Width / 60
'
'
'    picFrame.Width = ImgWidth
'    imgFrame.Width = Width
'
'    Dim i%
'    For i = 0 To 2
'        mClipControl(i).Left = Width - 255 * (3.5 - i)
'    Next i
    
    If Not bBitmapFound Then
        Call RePaint
    End If
End Sub

