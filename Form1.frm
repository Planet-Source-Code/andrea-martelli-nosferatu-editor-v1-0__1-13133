VERSION 5.00
Object = "{8A343B10-221C-4148-92FD-FA2ECA1C9C4E}#1.0#0"; "ETSLINKLABEL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nosferatu Editor v1.0"
   ClientHeight    =   4548
   ClientLeft      =   120
   ClientTop       =   396
   ClientWidth     =   5736
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4548
   ScaleWidth      =   5736
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MainPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      DrawStyle       =   3  'Dash-Dot
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3936
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   326
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   458
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   5520
      Begin VB.PictureBox destPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   12
         Left            =   1080
         ScaleHeight     =   1
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   12
      End
      Begin VB.PictureBox picVuoto 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   252
         Left            =   0
         Picture         =   "Form1.frx":BD76
         ScaleHeight     =   252
         ScaleWidth      =   12
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   12
      End
      Begin VB.Shape Shape1 
         Height          =   492
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   12
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   3  'Dot
         DrawMode        =   1  'Blackness
         Height          =   756
         Left            =   4440
         Top             =   2520
         Visible         =   0   'False
         Width           =   12
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Selezione"
      Height          =   372
      Left            =   4440
      TabIndex        =   0
      Top             =   4080
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "Download this code and many other visiting:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   4092
   End
   Begin EttsLinkLabel.EtsLinkLabel EtsLinkLabel1 
      Height          =   204
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   360
      Caption         =   "Visual Basic Italia"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Visual Basic Italia"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropColorAngleLeft=   15
      DropColorAngleTop=   15
      URL             =   "http://digilander.iol.it/VBItalia/INDICE.html"
   End
   Begin VB.Menu ModificaMainMnu 
      Caption         =   "Edit"
      Begin VB.Menu TagliaMnu 
         Caption         =   "&Cut"
         Enabled         =   0   'False
         Shortcut        =   ^T
      End
      Begin VB.Menu CopiaMnu 
         Caption         =   "&Copy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu IncollaMnu 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^I
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************'
'*     Nosferatu Editor v1.0  *'
'******************************'
'*      di Andrea Martelli    *'
'******************************'
'*                            *'
'* per scaricare gratuitamente*'
'* le guide utilizzate per    *'
'* sviluppare questo editor   *'
'* fare riferimento al sito:  *'
'******************************'
'http://digilander.iol.it/VBItalia/INDICE.html

Dim XX As Double, YY As Double
Dim XX1 As Double, YY1 As Double
Dim Xcopia As Double
Dim OkSeleziona As Boolean
Dim retval As Variant
Dim destL As Long
Dim Incolla As Boolean
Dim destT As Long
Dim destW As Long
Dim destH As Long
Dim Ycopia As Double
Dim Sposta As Boolean
Option Explicit
Private Sub CopiaPerSpostare()
'cancella l'immagine eventualmente
'creata in precedenza nella picturebox
'di destinazione
destPic.Cls
'richiama la funzione BitBlt.
retval = BitBlt(destPic.hDC, 0, 0, Shape1.Width, Shape1.Height, MainPic.hDC, XX, YY, SRCCOPY)
'aggiorna la picturebox di destinazione
destPic.Refresh
'Notare che alla lunghezza e all'altezza della
'Shape1 sono stati tolti due pixels mentre alle
'coordinate x ed y uno. Questo allo scopo di non
'copiare anche le linee della Shape1, ma solo quello che
'contiene
Shape1.Left = destPic.Left - 1
Shape1.Top = destPic.Top - 1
Shape1.Width = destPic.Width + 2
Shape1.Height = destPic.Height + 2
Shape1.BorderStyle = 1
'imposta il puntatore "size all" della
'selezione sulla quale si può operare
destPic.MousePointer = 15
'indica che si può operare sulla
'selezione
Sposta = True
End Sub
Private Sub Command1_Click()
'abilita i tasti taglia, copia, incolla
CopiaMnu.Enabled = True
TagliaMnu.Enabled = True
IncollaMnu.Enabled = True
'indica che si può
'lavorare sulla selezione
OkSeleziona = True
End Sub
Private Sub CopiaMnu_Click()
'elimina il precedente contenuto della
'clipboard
Clipboard.Clear
'mette l'immagine di destinazione nella
'clipboard
Clipboard.SetData destPic.Image
Exit Sub
End Sub

Private Sub destPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Sposta = True Then
'permette di muovere l'immagine di
'destinazione selezionata
If Button = vbLeftButton Then
'indica che si può spostare l'immagine
Xcopia = X
Ycopia = Y
End If
End If
End Sub
Private Sub destPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Sposta = True Then
If Button = vbLeftButton Then
Shape1.Visible = False
'sposta la bmp di destinazione in base
'ai movimenti del mouse su di essa
If X < Xcopia Then
destPic.Left = destPic.Left - (Xcopia - X)
End If
If X > Xcopia Then
destPic.Left = destPic.Left + (X - Xcopia)
End If
If Y < Ycopia Then
destPic.Top = destPic.Top - (Ycopia - Y)
End If
If Y > Ycopia Then
destPic.Top = destPic.Top + (Y - Ycopia)
End If
destL = destPic.Left
destT = destPic.Top
destW = destPic.Width
destH = destPic.Height
destPic.Refresh
End If
End If
End Sub
Private Sub destPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Sposta = True Then
Shape1.Visible = True
Shape1.Left = destPic.Left - 1
Shape1.Top = destPic.Top - 1
Shape1.Width = destPic.Width + 2
Shape1.Height = destPic.Height + 2
End If
End Sub

Private Sub IncollaMnu_Click()
destPic.Visible = True
destPic.AutoSize = True
destPic.Picture = Clipboard.GetData
destPic.Left = 0
destPic.Top = 0
Shape1.Visible = True
Shape1.Left = destPic.Left - 1
Shape1.Top = destPic.Top - 1
Shape1.Width = destPic.Width + 2
Shape1.Height = destPic.Height + 2
Incolla = True
Sposta = True
End Sub

Private Sub MainPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'se è stato premuto il tasto di selezione
If OkSeleziona = True Then
'se è possibile spostare l'immagine
If Sposta = True Then
'copia l'immagine nella posizione dove
'si trova nel momento in cui si clicca
'sulla picturebox principale
If Incolla = False Then
retval = BitBlt(MainPic.hDC, destL, destT, Shape1.Width, Shape1.Height, MainPic.hDC, XX, YY, SRCCOPY)
Else
MainPic.PaintPicture destPic, destPic.Left, destPic.Top
Incolla = False
End If
'non si può più spostare la selezione
'che diventa ormai parte della picture
'principale
Sposta = False
End If
'memorizza le coordinate del punto in
'cui inizia la selezione ossia il vertice
'in alto a sx del rettangolo che si
'forma cliccando sulla picturebox principale
'e trascinando il mouse
XX = X: YY = Y
'posizione la shape2 e nasconde la
'shape1
Shape1.Visible = False
Shape2.Shape = 0
Shape2.Visible = True
Shape2.Left = X: Shape2.Top = Y
Shape2.Width = 0: Shape2.Height = 0
End If
End Sub
Private Sub MainPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'se è stato premuto il tasto di selezione
If OkSeleziona = True Then
Shape2.Left = IIf(X > XX, XX, X)
Shape2.Top = IIf(Y > YY, YY, Y)
Shape2.Height = Abs(Y - YY)
Shape2.Width = Abs(X - XX)
End If
End Sub
Private Sub MainPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If OkSeleziona = True Then
Shape1.Visible = True
Shape1.Left = Shape2.Left
Shape1.Top = Shape2.Top
Shape1.Height = Shape2.Height
Shape1.Width = Shape2.Width
Shape2.Visible = False
XX1 = X
YY1 = Y
destPic.Top = Shape2.Top
destPic.Left = Shape2.Left
destPic.Height = Shape2.Height
destPic.Width = Shape2.Width
destPic.Visible = True
Call CopiaPerSpostare
End If
End Sub

Private Sub TagliaMnu_Click()
'indica che si può procedere ad
'immagazzinare in memoria l'immagine
'sorgente
If Sposta = True Then
'necessaria ripetizione in quanto l'utente
'potrebbe tagliare la selezione senza
'aver prima mosso l'immagine
destL = destPic.Left
destT = destPic.Top
destW = destPic.Width
destH = destPic.Height
'mette l'immagine in memoria nella
'Clipboard
Clipboard.Clear
Clipboard.SetData destPic.Image
MainPic.PaintPicture picVuoto.Picture, destL, destT, Abs(XX1 - XX), Abs(YY1 - YY)
retval = BitBlt(destPic.hDC, 0, 0, destPic.Width, destPic.Height, MainPic.hDC, destL, destT, SRCCOPY)
Shape1.Visible = False
End If
End Sub

