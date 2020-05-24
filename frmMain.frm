VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "MBMenschAergereDichNicht2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MBMenschAergereDichNicht2"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows-Standard
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

