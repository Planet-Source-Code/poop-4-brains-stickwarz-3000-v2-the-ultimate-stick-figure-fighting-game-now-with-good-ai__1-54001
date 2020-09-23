VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StickWarz 3000 by poop4brains (Kevin Fleet)"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   286
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   353
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Options 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   0
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   351
      TabIndex        =   60
      Top             =   360
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Timer tmrFPS 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2520
         Top             =   240
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   101
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtFPS 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   100
         Text            =   "0"
         Top             =   1680
         Width           =   735
      End
      Begin VB.ComboBox cmbKey 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   15
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   98
         Top             =   2760
         Width           =   615
      End
      Begin VB.ComboBox cmbKey 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   14
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   2760
         Width           =   615
      End
      Begin VB.ComboBox cmbKey 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   13
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   96
         Top             =   2760
         Width           =   615
      End
      Begin VB.ComboBox cmbKey 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   2760
         Width           =   615
      End
      Begin VB.ComboBox cmbKey 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   2760
         Width           =   615
      End
      Begin VB.ComboBox cmbKey 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   2760
         Width           =   615
      End
      Begin VB.ComboBox cmbKey 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Top             =   2760
         Width           =   615
      End
      Begin VB.ComboBox cmbKey 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   91
         Top             =   2760
         Width           =   615
      End
      Begin VB.ComboBox cmbKey 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   90
         Top             =   2400
         Width           =   615
      End
      Begin VB.ComboBox cmbKey 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   2400
         Width           =   615
      End
      Begin VB.ComboBox cmbKey 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   2400
         Width           =   615
      End
      Begin VB.ComboBox cmbKey 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   2400
         Width           =   615
      End
      Begin VB.ComboBox cmbKey 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   2400
         Width           =   615
      End
      Begin VB.ComboBox cmbKey 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   2400
         Width           =   615
      End
      Begin VB.ComboBox cmbKey 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   2400
         Width           =   615
      End
      Begin VB.ComboBox cmbKey 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   73
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   72
         Top             =   3240
         Width           =   1095
      End
      Begin VB.HScrollBar hsSpeed 
         Height          =   255
         LargeChange     =   1000
         Left            =   600
         Max             =   10000
         Min             =   500
         SmallChange     =   100
         TabIndex        =   63
         Top             =   1200
         Value           =   500
         Width           =   3735
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   35
         Left            =   3720
         TabIndex        =   102
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "FPS(60-70 optimal):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   34
         Left            =   1800
         TabIndex        =   99
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   33
         Left            =   4560
         TabIndex        =   89
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Throw"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   32
         Left            =   3960
         TabIndex        =   88
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Blast"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   3360
         TabIndex        =   87
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Kick"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   30
         Left            =   2760
         TabIndex        =   86
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Punch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   29
         Left            =   2160
         TabIndex        =   83
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Right"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   28
         Left            =   1560
         TabIndex        =   82
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Left"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   960
         TabIndex        =   81
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Jump"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   360
         TabIndex        =   75
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "P2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   74
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "P1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   71
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Slower Computer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   600
         TabIndex        =   69
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Faster Computer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   2760
         TabIndex        =   68
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Keys"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   480
         TabIndex        =   66
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Faster"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   600
         TabIndex        =   65
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Slower"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   3000
         TabIndex        =   64
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Speed Control"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   480
         TabIndex        =   62
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   240
         TabIndex        =   61
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   4320
      TabIndex        =   109
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   120
      TabIndex        =   108
      Top             =   5400
      Width           =   4095
   End
   Begin VB.TextBox txtChat 
      Height          =   975
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   107
      Top             =   4320
      Width           =   5055
   End
   Begin VB.PictureBox EM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   960
      Picture         =   "frmGame.frx":030A
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   70
      Top             =   6000
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.PictureBox ES 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   1080
      Picture         =   "frmGame.frx":634C
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   67
      Top             =   5880
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.PictureBox NewGame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      DrawWidth       =   5
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3630
      Left            =   240
      Picture         =   "frmGame.frx":C38E
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   29
      Top             =   360
      Visible         =   0   'False
      Width           =   4830
      Begin MSWinsockLib.Winsock wsHost2 
         Left            =   3720
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   12700
      End
      Begin MSWinsockLib.Winsock wsHost 
         Left            =   3840
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsClient 
         Left            =   3360
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   12700
      End
      Begin VB.CommandButton cmdHost 
         Caption         =   "Host"
         Height          =   255
         Left            =   3720
         TabIndex        =   105
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel1 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   103
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">>"
         Height          =   255
         Left            =   600
         TabIndex        =   56
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "<<"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "Start!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   54
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CheckBox chkAllowChange 
         BackColor       =   &H00FF8080&
         Caption         =   "Allow Character Change"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   53
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CheckBox chkMarker 
         BackColor       =   &H00FF8080&
         Caption         =   "Show Markers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   52
         Top             =   3120
         Width           =   1095
      End
      Begin VB.ComboBox cmbControl 
         Height          =   315
         Index           =   1
         ItemData        =   "frmGame.frx":447D0
         Left            =   3480
         List            =   "frmGame.frx":447DA
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   2520
         Width           =   1215
      End
      Begin VB.ComboBox cmbControl 
         Height          =   315
         Index           =   0
         ItemData        =   "frmGame.frx":447EF
         Left            =   3480
         List            =   "frmGame.frx":447F9
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   2040
         Width           =   1215
      End
      Begin VB.ComboBox cmbChar 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         ItemData        =   "frmGame.frx":4480E
         Left            =   2040
         List            =   "frmGame.frx":4481B
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   2520
         Width           =   1335
      End
      Begin VB.ComboBox cmbChar 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         ItemData        =   "frmGame.frx":44842
         Left            =   2040
         List            =   "frmGame.frx":4484F
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   960
         MaxLength       =   8
         TabIndex        =   44
         Text            =   "Player 2"
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   960
         MaxLength       =   8
         TabIndex        =   43
         Text            =   "Player 1"
         Top             =   2040
         Width           =   975
      End
      Begin VB.PictureBox MapPre 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   120
         Picture         =   "frmGame.frx":44876
         ScaleHeight     =   63
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   63
         TabIndex        =   30
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdJoin 
         Caption         =   "Join"
         Height          =   255
         Left            =   2760
         TabIndex        =   106
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Control"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   3480
         TabIndex        =   49
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Character"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   2040
         TabIndex        =   46
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   960
         TabIndex        =   45
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   42
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   41
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   8
         X2              =   312
         Y1              =   112
         Y2              =   112
      End
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "<None>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   40
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblAuthor 
         BackStyle       =   0  'Transparent
         Caption         =   "<None>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   39
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblDifficulty 
         BackStyle       =   0  'Transparent
         Caption         =   "<None>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   38
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "<None>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   37
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "<None>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   36
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   1200
         TabIndex        =   35
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Creation Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   1200
         TabIndex        =   34
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Author:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   1200
         TabIndex        =   33
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Difficulty:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   1200
         TabIndex        =   32
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Map Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   1200
         TabIndex        =   31
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.ListBox lstMapInfo 
      Height          =   2205
      Left            =   5640
      TabIndex        =   57
      Top             =   240
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.PictureBox SS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5280
      Index           =   2
      Left            =   7440
      Picture         =   "frmGame.frx":478B8
      ScaleHeight     =   352
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   24
      Top             =   5880
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox SM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5280
      Index           =   2
      Left            =   7440
      Picture         =   "frmGame.frx":580FA
      ScaleHeight     =   352
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   23
      Top             =   600
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox Logo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   240
      Picture         =   "frmGame.frx":6893C
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   319
      TabIndex        =   22
      Top             =   360
      Width           =   4815
      Begin VB.PictureBox pcLMap 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   3120
         ScaleHeight     =   105
         ScaleWidth      =   1545
         TabIndex        =   58
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading Map Info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   3120
         TabIndex        =   59
         Top             =   3120
         Width           =   1575
      End
   End
   Begin VB.PictureBox SM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5280
      Index           =   1
      Left            =   6480
      Picture         =   "frmGame.frx":6AFE8
      ScaleHeight     =   352
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   21
      Top             =   600
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox SS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5280
      Index           =   1
      Left            =   6480
      Picture         =   "frmGame.frx":7B82A
      ScaleHeight     =   352
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   20
      Top             =   5880
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox FM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6120
      Picture         =   "frmGame.frx":8C06C
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   7
      Top             =   4920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox FS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6360
      Picture         =   "frmGame.frx":8C3AE
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   6
      Top             =   4920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox GS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   5640
      Picture         =   "frmGame.frx":8C6F0
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox GM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   5640
      Picture         =   "frmGame.frx":8CA32
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox bSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   360
      Picture         =   "frmGame.frx":8CD74
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   3
      Top             =   5880
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.PictureBox bMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   360
      Picture         =   "frmGame.frx":C51B6
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   2
      Top             =   6120
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.PictureBox SM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5280
      Index           =   0
      Left            =   5520
      Picture         =   "frmGame.frx":FD5F8
      ScaleHeight     =   352
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox SS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5280
      Index           =   0
      Left            =   5520
      Picture         =   "frmGame.frx":10DE3A
      ScaleHeight     =   352
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox Menu 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   240
      Picture         =   "frmGame.frx":11E67C
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   319
      TabIndex        =   25
      Top             =   360
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Label cmdExit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   28
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label cmdOptions 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   27
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label cmdNewMenu 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "New Game"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   26
         Top             =   2520
         Width           =   1215
      End
   End
   Begin VB.PictureBox Frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   4125
      Left            =   120
      ScaleHeight     =   273
      ScaleMode       =   0  'User
      ScaleWidth      =   343.045
      TabIndex        =   8
      Top             =   120
      Width           =   5100
      Begin VB.PictureBox pEN 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   1
         Left            =   2760
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   55
         TabIndex        =   18
         Top             =   3915
         Width           =   855
      End
      Begin VB.PictureBox pEN 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   0
         Left            =   2760
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   55
         TabIndex        =   17
         Top             =   45
         Width           =   855
      End
      Begin VB.PictureBox pHP 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H008080FF&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   1
         Left            =   1320
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   55
         TabIndex        =   15
         Top             =   3915
         Width           =   855
      End
      Begin VB.PictureBox pHP 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H008080FF&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   0
         Left            =   1320
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   55
         TabIndex        =   14
         Top             =   45
         Width           =   855
      End
      Begin VB.PictureBox Board 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         DrawWidth       =   5
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3630
         Left            =   120
         Picture         =   "frmGame.frx":156ABE
         ScaleHeight     =   240
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   320
         TabIndex        =   9
         Top             =   240
         Width           =   4830
         Begin VB.Label lblQuit 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Quit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   104
            Top             =   2640
            Visible         =   0   'False
            Width           =   4815
         End
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "EN:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   2400
         TabIndex        =   19
         Top             =   3870
         Width           =   375
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "EN:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   16
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   13
         Top             =   3870
         Width           =   375
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   12
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   3870
         Width           =   735
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   0
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tSpeed As Long, tmpSpeed As Long
Dim MapI As Long
Dim FPS(1 To 5) As Long, sFPS As Long, iFPS As Long

Function MainLoop()
Dim C As Long, CC As Long
Running = True
Do Until Running = False
    If C >= tSpeed Then
        If Paused = False Then
            Board.Cls
            BitBlt Board.hDC, 0, 0, 320, 240, bMask.hDC, 0, 0, vbSrcAnd
            BitBlt Board.hDC, 0, 0, 320, 240, bSprite.hDC, 0, 0, vbSrcInvert
            For CC = 1 To 2
                If wsClient.State = sckClosed And wsHost.State = sckClosed Then
                    GetKeys P(CC), P(CC).Moves(1), P(CC).Moves(2), P(CC).Moves(3), P(CC).Moves(4), P(CC).Moves(5), P(CC).Moves(6), P(CC).Moves(7), P(CC).Moves(8)
                ElseIf wsClient.State = sckConnected And CC = 2 Then
                    GetKeys P(CC), P(CC).Moves(1), P(CC).Moves(2), P(CC).Moves(3), P(CC).Moves(4), P(CC).Moves(5), P(CC).Moves(6), P(CC).Moves(7), P(CC).Moves(8)
                ElseIf wsHost.State = sckConnected And CC = 1 Then
                    GetKeys P(CC), P(CC).Moves(1), P(CC).Moves(2), P(CC).Moves(3), P(CC).Moves(4), P(CC).Moves(5), P(CC).Moves(6), P(CC).Moves(7), P(CC).Moves(8)
                End If
            Next CC
            MovePlayers
            MoveShots
            MoveExplo
            If CheckWinner <> 0 Then
                Board.DrawWidth = 1
                Dim X As Long, Y As Long, msg As String
                For X = 0 To 320 Step 2
                    For Y = 0 To 240 Step 2
                        Board.PSet (X, Y), &H808080
                    Next Y
                Next X
                Running = False
                dSound.PlaySound 3, False
                If CheckWinner = 3 Then
                    msg = "The match was a tie!"
                Else
                    msg = P(CheckWinner).Name & " has won!"
                End If
                Board.CurrentX = Board.ScaleWidth \ 2 - Board.TextWidth(msg) \ 2
                Board.CurrentY = Board.ScaleHeight \ 2 - Board.TextHeight("|") \ 2
                Board.Print msg
            End If
            C = 0
            sFPS = sFPS + 1
        End If
    Else
        C = C + 1
    End If
    DoEvents
Loop
End Function

Private Sub Board_Click()
dSound.PlaySound 1, False
If Running = True Then
    Paused = IIf(Paused = True, False, True)
    If Paused = True Then
        Board.ForeColor = vbWhite
        Board.DrawWidth = 1
        Dim X As Long, Y As Long
        For X = 0 To 320 Step 2
            For Y = 0 To 240 Step 2
                Board.PSet (X, Y), &H808080
            Next Y
        Next X
        Board.CurrentX = Board.ScaleWidth \ 2 - Board.TextWidth("Paused") \ 2
        Board.CurrentY = Board.ScaleHeight \ 2 - Board.TextHeight("|") \ 2
        Board.Print "Paused"
        lblQuit.Visible = True
    Else
        lblQuit.Visible = False
    End If
Else
    Logo_Click
End If
End Sub

Private Sub cmdBack_Click()
Dim I As Long, I2 As Long, MapName As String
dSound.PlaySound 1, False
MapI = MapI - 1: If MapI < 1 Then MapI = lstMapInfo.ListCount / 6
For I = 0 To lstMapInfo.ListCount - 1 Step 6
    I2 = I2 + 1
    If I2 = MapI Then
        lblName.Caption = lstMapInfo.List(I)
        lblDifficulty.Caption = Split(lstMapInfo.List(I + 1), "|")(0)
        P(1).X = Val(Split(Split(lstMapInfo.List(I + 1), "|")(1), ",")(0))
        P(1).Y = Val(Split(Split(lstMapInfo.List(I + 1), "|")(1), ",")(1))
        P(2).X = Val(Split(Split(lstMapInfo.List(I + 1), "|")(2), ",")(0))
        P(2).Y = Val(Split(Split(lstMapInfo.List(I + 1), "|")(2), ",")(1))
        lblAuthor.Caption = lstMapInfo.List(I + 2)
        lblDate.Caption = lstMapInfo.List(I + 3)
        lblDescription.Caption = lstMapInfo.List(I + 4)
        MapName = lstMapInfo.List(I + 5)
        MapPre.Picture = LoadPicture(App.Path & "\maps\" & MapName & "_icon.bmp")
        bMask.Picture = LoadPicture(App.Path & "\maps\" & MapName & "_mask.bmp")
        bSprite.Picture = LoadPicture(App.Path & "\maps\" & MapName & "_sprite.bmp")
        Exit For
    End If
Next I
End Sub

Private Sub cmdCancel_Click()
dSound.PlaySound 1, False
Options.Visible = False
Menu.Visible = True
End Sub

Private Sub cmdCancel1_Click()
If wsHost.State <> sckClosed Then wsHost.Close
If wsClient.State <> sckClosed Then wsClient.Close
EnableConsole
Me.Height = 4680
dSound.PlaySound 1, False
NewGame.Visible = False
Menu.Visible = True
End Sub

Private Sub cmdExit_Click()
dSound.PlaySound 1, False
Unload Me
End Sub

Private Sub cmdHost_Click()
Dim Port As Long
Port = Val(InputBox("On what port do you want to connect?", "Port", "12700"))
If Port <= 0 Then Exit Sub
cmdJoin.Enabled = False
cmdHost.Enabled = False
wsHost2.LocalPort = Port
wsHost2.Listen
Me.Height = 6165
End Sub

Private Sub cmdJoin_Click()
Dim IP As String, Port As Long
IP = InputBox("What IP do you want to connect to?", "IP", "127.0.0.1")
If IP = "" Then Exit Sub
Port = Val(InputBox("On what port do you want to connect?", "Port", "12700"))
If Port <= 0 Then Exit Sub
cmdJoin.Enabled = False
cmdHost.Enabled = False
wsClient.Connect IP, Port
Me.Height = 6165
End Sub

Private Sub cmdNew_Click()
Dim I As Long, I2 As Long, MapName As String

If lblName.Caption = "<None>" Then
    MsgBox "Please select a map!", vbCritical, "Error"
    Exit Sub
End If
If cmbChar(0).ListIndex < 0 Or cmbChar(1).ListIndex < 0 Then
    MsgBox "Please select a character!", vbCritical, "Error"
    Exit Sub
End If
If txtName(0).Text = "" Or txtName(0).Text = "" Then
    MsgBox "Please fill in a name!", vbCritical, "Error"
    Exit Sub
End If

For I = 0 To lstMapInfo.ListCount - 1 Step 6
    I2 = I2 + 1
    If I2 = MapI Then
        lblName.Caption = lstMapInfo.List(I)
        lblDifficulty.Caption = Split(lstMapInfo.List(I + 1), "|")(0)
        P(1).X = Val(Split(Split(lstMapInfo.List(I + 1), "|")(1), ",")(0))
        P(1).Y = Val(Split(Split(lstMapInfo.List(I + 1), "|")(1), ",")(1))
        P(2).X = Val(Split(Split(lstMapInfo.List(I + 1), "|")(2), ",")(0))
        P(2).Y = Val(Split(Split(lstMapInfo.List(I + 1), "|")(2), ",")(1))
        lblAuthor.Caption = lstMapInfo.List(I + 2)
        lblDate.Caption = lstMapInfo.List(I + 3)
        lblDescription.Caption = lstMapInfo.List(I + 4)
        MapName = lstMapInfo.List(I + 5)
        MapPre.Picture = LoadPicture(App.Path & "\maps\" & MapName & "_icon.bmp")
        bMask.Picture = LoadPicture(App.Path & "\maps\" & MapName & "_mask.bmp")
        bSprite.Picture = LoadPicture(App.Path & "\maps\" & MapName & "_sprite.bmp")
        Board.Picture = LoadPicture(App.Path & "\maps\" & MapName & "_back.bmp")
    End If
Next I

Board.Visible = True
lblQuit.Visible = False
mciSendString "stop MFile", 0&, 0, 0
mciSendString "close MFile", 0&, 0, 0
dSound.PlaySound 1, False
ShowMarker = IIf(chkMarker.Value = vbChecked, True, False)
LetChange = IIf(chkAllowChange.Value = vbChecked, True, False)
Menu.Visible = False
NewGame.Visible = False
ClearMem
NewPlayer P(1), P(1).X, P(1).Y, 1, cmbChar(0).ListIndex, 500, txtName(0).Text, cmbControl(0).ListIndex
NewPlayer P(2), P(2).X, P(2).Y, 2, cmbChar(1).ListIndex, 500, txtName(1).Text, cmbControl(1).ListIndex
lblLabel(0).Caption = P(1).Name
lblLabel(1).Caption = P(2).Name
MainLoop
End Sub

Private Sub cmdNewMenu_Click()
dSound.PlaySound 1, False
NewGame.Visible = True
End Sub

Private Sub cmdNext_Click()
Dim I As Long, I2 As Long, MapName As String
dSound.PlaySound 1, False
MapI = MapI + 1: If MapI > lstMapInfo.ListCount / 6 Then MapI = 1
For I = 0 To lstMapInfo.ListCount - 1 Step 6
    I2 = I2 + 1
    If I2 = MapI Then
        lblName.Caption = lstMapInfo.List(I)
        lblDifficulty.Caption = Split(lstMapInfo.List(I + 1), "|")(0)
        P(1).X = Val(Split(Split(lstMapInfo.List(I + 1), "|")(1), ",")(0))
        P(1).Y = Val(Split(Split(lstMapInfo.List(I + 1), "|")(1), ",")(1))
        P(2).X = Val(Split(Split(lstMapInfo.List(I + 1), "|")(2), ",")(0))
        P(2).Y = Val(Split(Split(lstMapInfo.List(I + 1), "|")(2), ",")(1))
        lblAuthor.Caption = lstMapInfo.List(I + 2)
        lblDate.Caption = lstMapInfo.List(I + 3)
        lblDescription.Caption = lstMapInfo.List(I + 4)
        MapName = lstMapInfo.List(I + 5)
        MapPre.Picture = LoadPicture(App.Path & "\maps\" & MapName & "_icon.bmp")
        bMask.Picture = LoadPicture(App.Path & "\maps\" & MapName & "_mask.bmp")
        bSprite.Picture = LoadPicture(App.Path & "\maps\" & MapName & "_sprite.bmp")
        Board.Picture = LoadPicture(App.Path & "\maps\" & MapName & "_back.bmp")
    End If
Next I
End Sub

Private Sub cmdOk_Click()
Dim CC As Long, CC2 As Long, cI As Long
For CC = 1 To 2
    For CC2 = 1 To 8
        P(CC).Moves(CC2) = cmbKey(cI).ItemData(cmbKey(cI).ListIndex)
        cI = cI + 1
    Next CC2
Next CC
tSpeed = hsSpeed.Value
dSound.PlaySound 1, False
Options.Visible = False
Menu.Visible = True
End Sub

Private Sub cmdOptions_Click()
dSound.PlaySound 1, False
Options.Visible = True
Menu.Visible = False
Logo.Visible = False
Board.Visible = False
End Sub

Private Sub cmdSend_Click()
txtSend.Text = Replace(txtSend.Text, "|¿|", "")
If wsHost.State = sckConnected Then wsHost.SendData "MSG|¿|" & txtSend.Text
If wsClient.State = sckConnected Then wsClient.SendData "MSG|¿|" & txtSend.Text
txtSend.Text = ""
End Sub

Private Sub cmdTest_Click()
tmpSpeed = tSpeed
tSpeed = hsSpeed.Value
iFPS = 0
sFPS = 0
tmrFPS.Enabled = True
cmdTest.Enabled = False
cmdNew_Click
End Sub

Private Sub Form_Load()
Dim T As String, D As String, A As String, Dt As String, DS As String, N As String, CC As Long, CC2 As Long, cI As Long, I As Long
Me.Visible = True
Open App.Path & "\maps\maps.dat" For Input As #1
Do Until EOF(1)
    Input #1, T, D, A, Dt, DS, N
    lstMapInfo.AddItem T
    lstMapInfo.AddItem D
    lstMapInfo.AddItem A
    lstMapInfo.AddItem Dt
    lstMapInfo.AddItem DS
    lstMapInfo.AddItem N
    DoEvents
Loop
Close #1
pcLMap.Line (0, 0)-(pcLMap.ScaleWidth \ 2, pcLMap.ScaleHeight), vbRed, BF
lblLabel(16).Caption = "Loading Preferences"
Open App.Path & "\maps\prefs.dat" For Input As #1
Input #1, tSpeed
hsSpeed.Value = tSpeed
For CC = 1 To 2
    For CC2 = 1 To 8
        Input #1, P(CC).Moves(CC2)
    Next CC2
Next CC
Close #1
Open App.Path & "\maps\keys.dat" For Input As #1
Do Until EOF(1)
    Input #1, T, D
    For CC = 0 To 15
        cmbKey(CC).AddItem T
        cmbKey(CC).ItemData(cmbKey(CC).ListCount - 1) = D
    Next CC
    DoEvents
Loop
Close #1
cI = 0
For CC = 1 To 2
    For CC2 = 1 To 8
        For I = 0 To cmbKey(cI).ListCount - 1
            If cmbKey(cI).ItemData(I) = P(CC).Moves(CC2) Then
                cmbKey(cI).ListIndex = I
                Exit For
            End If
        Next I
        cI = cI + 1
    Next CC2
Next CC

pcLMap.Line (0, 0)-(pcLMap.ScaleWidth, pcLMap.ScaleHeight), vbRed, BF
lblLabel(16).Caption = "Done, Click to start"
Randomize
dSound.Initialize_Engine Me.Hwnd
dSound.LoadWavToChannel 1, App.Path & "\sounds\pop.wav"
dSound.SetVolume 1, 10
dSound.LoadWavToChannel 2, App.Path & "\sounds\explo.wav"
dSound.SetVolume 2, 10
dSound.LoadWavToChannel 3, App.Path & "\sounds\success.wav"
dSound.SetVolume 3, 10
dSound.LoadWavToChannel 4, App.Path & "\sounds\ow.wav"
dSound.SetVolume 4, 10
dSound.LoadWavToChannel 5, App.Path & "\sounds\poop.wav"
dSound.SetVolume 5, 10
dSound.LoadWavToChannel 6, App.Path & "\sounds\eball.wav"
dSound.SetVolume 6, 10
cmbChar(0).ListIndex = 0
cmbChar(1).ListIndex = 0
cmbControl(0).ListIndex = 0
cmbControl(1).ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim CC As Long, CC2 As Long
Running = False
DoEvents
Call mciSendString("Close All", 0&, 0, 0)
dSound.Terminate_Engine
Kill App.Path & "\maps\prefs.dat"
Open App.Path & "\maps\prefs.dat" For Output As #1
Write #1, tSpeed
For CC = 1 To 2
    For CC2 = 1 To 8
        Write #1, P(CC).Moves(CC2)
    Next CC2
Next CC
Close #1
If wsClient.State <> sckClosed Then wsClient.Close
If wsHost.State <> sckClosed Then wsHost.Close
End Sub

Private Sub hsSpeed_Change()
lblLabel(35).Caption = hsSpeed.Value
End Sub

Private Sub lblQuit_Click()
P(1).Act = False
P(2).Act = False
Paused = False
End Sub

Private Sub Logo_Click()
Dim MIDIPATH As String
Logo.Visible = False
Menu.Visible = True
Options.Visible = False
Board.Visible = False
MIDIPATH = GetShortFileName(App.Path & "\sounds\rock on.mid")
mciSendString "open " & MIDIPATH & " Type sequencer Alias MFile", 0&, 0, 0
mciSendString "play MFile", 0&, 0, 0
End Sub

Private Sub tmrFPS_Timer()
iFPS = iFPS + 1
FPS(iFPS) = sFPS
sFPS = 0
If iFPS >= 5 Then
    txtFPS.Text = (FPS(1) + FPS(2) + FPS(3) + FPS(4) + FPS(5)) \ 5
    Running = False
    tSpeed = tmpSpeed
    cmdTest.Enabled = True
    tmrFPS.Enabled = False
End If
End Sub

Function EnableConsole()
cmdNext.Enabled = True
cmdBack.Enabled = True
cmdJoin.Enabled = True
cmdHost.Enabled = True
txtName(0).Enabled = True
txtName(1).Enabled = True
cmbChar(0).Enabled = True
cmbChar(1).Enabled = True
cmbControl(0).Enabled = True
cmbControl(1).Enabled = True
chkMarker.Enabled = True
chkAllowChange.Enabled = True
cmdNew.Enabled = False
End Function

Function Chat(msg)
txtChat.Text = txtChat.Text & msg & vbCrLf
End Function

Private Sub wsClient_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
wsClient.GetData Data
Select Case UCase(Split(Data, "|¿|")(0))
Case "MSG"
    Chat Split(Data, "|¿|")(1)
Case "PLY"
    P(2).X = Val(Split(Split(Data, "|¿|")(1), ",")(0))
    P(2).Y = Val(Split(Split(Data, "|¿|")(1), ",")(1))
    DoKeys P(2), Split(Data, "|¿|")(2)
End Select
End Sub

Private Sub wsHost_Connect()
cmbControl(0).Enabled = False
cmbControl(1).Enabled = False
cmbControl(0).ListIndex = 0
cmbControl(1).ListIndex = 0
txtName(0).Enabled = False
cmbChar(0).Enabled = False
cmdNext.Enabled = False
cmdBack.Enabled = False
Chat "Connection Established"
End Sub

Private Sub wsHost_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
wsHost.GetData Data
Select Case UCase(Split(Data, "|¿|")(0))
Case "MSG"
    Chat Split(Data, "|¿|")(1)
Case "PLY"
    P(2).X = Val(Split(Split(Data, "|¿|")(1), ",")(0))
    P(2).Y = Val(Split(Split(Data, "|¿|")(1), ",")(1))
    DoKeys P(2), Split(Data, "|¿|")(2)
End Select
End Sub

Private Sub wsHost_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
cmbControl(0).Enabled = False
cmbControl(1).Enabled = False
cmbControl(0).ListIndex = 0
cmbControl(1).ListIndex = 0
txtName(1).Enabled = False
cmbChar(1).Enabled = False
Chat "Connection established"
End Sub

Private Sub wsHost2_ConnectionRequest(ByVal requestID As Long)
wsHost.Accept requestID
End Sub

