VERSION 5.00
Begin VB.Form frmMS 
   Caption         =   "Masker/Spriter"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   ScaleHeight     =   484
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   603
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "blank"
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdMask 
      Caption         =   "Do It!"
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox Board 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   215
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   279
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
End
Attribute VB_Name = "frmMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMask_Click()
Board.Picture = LoadPicture(App.Path & "\" & txtPath.Text & ".bmp")
Dim X, Y, C
C = Board.Point(0, 0)
For X = 0 To Board.ScaleWidth
    For Y = 0 To Board.ScaleHeight
        If Board.Point(X, Y) = C Then
            Board.PSet (X, Y), vbBlack
        Else
            Board.PSet (X, Y), Board.Point(X, Y)
        End If
    Next Y
Next X
SavePicture Board.Image, App.Path & "\" & txtPath.Text & "_sprite.bmp"

Board.Cls
Board.Picture = LoadPicture(App.Path & "\" & txtPath.Text & ".bmp")
C = Board.Point(0, 0)
For X = 0 To Board.ScaleWidth
    For Y = 0 To Board.ScaleHeight
        If Board.Point(X, Y) = C Then
            Board.PSet (X, Y), vbWhite
        Else
            Board.PSet (X, Y), vbBlack
        End If
    Next Y
Next X
SavePicture Board.Image, App.Path & "\" & txtPath.Text & "_mask.bmp"
End Sub
