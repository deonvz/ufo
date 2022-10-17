VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5385
   ClientLeft      =   660
   ClientTop       =   600
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5385
   ScaleWidth      =   3840
   Begin VB.Timer tmrMove 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1320
      Top             =   2880
   End
   Begin VB.PictureBox picHidden 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5325
      Left            =   1800
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   355
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   251
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   3765
   End
   Begin VB.PictureBox picXMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   2760
      Picture         =   "Form1.frx":8301
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picX 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   2760
      Picture         =   "Form1.frx":87E3
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5385
      Left            =   0
      Picture         =   "Form1.frx":8CC5
      ScaleHeight     =   355
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   251
      TabIndex        =   2
      Top             =   0
      Width           =   3825
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deon van Zyl
Option Explicit

Private Const MERGEPAINT = &HBB0226
Private Const SRCAND = &H8800C6
Private Const SRCCOPY = &HCC0020
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

' Variables for positioning the image.
Private OldX As Single
Private OldY As Single
Private CurX As Single
Private CurY As Single
Private PicWid As Single
Private PicHgt As Single
Private Xmax As Single
Private Ymax As Single
Private NewX As Single
Private NewY As Single
Private Dx As Single
Private Dy As Single
Private DistToMove As Single

Private Const MOVE_OFFSET = 10

' Draw the picture at (CurX, CurY).
Private Sub DrawPicture()
    ' Fix the part of the image that was covered.
    BitBlt picCanvas.hDC, _
        OldX, OldY, PicWid, PicHgt, _
        picHidden.hDC, OldX, OldY, SRCCOPY
    OldX = CurX
    OldY = CurY

    ' Paint on the new image.
    BitBlt picCanvas.hDC, _
        CurX, CurY, PicWid, PicHgt, _
        picXMask.hDC, 0, 0, MERGEPAINT
    BitBlt picCanvas.hDC, _
        CurX, CurY, PicWid, PicHgt, _
        picX.hDC, 0, 0, SRCAND

    ' Update the display.
    picCanvas.Refresh
End Sub

' Save picCanvas's original bitmap bytes,
' initialize values, and draw the initial picture.
Private Sub Form_Load()
    ' Make the form fit the picture.
    Width = (Width - ScaleWidth) + picCanvas.Width
    Height = (Height - ScaleHeight) + picCanvas.Height

    PicWid = picX.ScaleWidth
    PicHgt = picX.ScaleHeight
    Xmax = picCanvas.ScaleWidth - PicWid
    Ymax = picCanvas.ScaleHeight - PicHgt
    OldX = 30
    OldY = 30
    CurX = 30
    CurY = 30

    DrawPicture
End Sub

' Stop dragging.
Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim dist As Single

    ' See where to move the image.
    NewX = x - PicWid / 2
    NewY = y - PicHgt / 2
    If NewX < 0 Then NewX = 0
    If NewX > Xmax Then NewX = Xmax
    If NewY < 0 Then NewY = 0
    If NewY > Ymax Then NewY = Ymax

    ' Calculate the moving offsets.
    Dx = NewX - CurX
    Dy = NewY - CurY
    DistToMove = Sqr(Dx * Dx + Dy * Dy)
    Dx = Dx / DistToMove * MOVE_OFFSET
    Dy = Dy / DistToMove * MOVE_OFFSET

    ' Enable the move timer.
    tmrMove.Enabled = True
End Sub

' Move the image closer to its destination.
Private Sub tmrMove_Timer()
    DistToMove = DistToMove - MOVE_OFFSET
    If DistToMove <= 0 Then
        CurX = NewX
        CurY = NewY
        tmrMove.Enabled = False
    Else
        CurX = CurX + Dx
        CurY = CurY + Dy
    End If

    DrawPicture
End Sub


