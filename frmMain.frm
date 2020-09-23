VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11490
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   766
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1022
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame turn1 
      BackColor       =   &H80000008&
      Height          =   2175
      Left            =   13200
      TabIndex        =   150
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
      Begin VB.Label lblTurn 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   120
         TabIndex        =   152
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Labels 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Player Turn:"
         BeginProperty Font 
            Name            =   "VAGRundschriftD"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   735
         Index           =   0
         Left            =   240
         TabIndex        =   151
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.PictureBox board1 
      BackColor       =   &H00404040&
      Height          =   9615
      Left            =   3120
      ScaleHeight     =   9555
      ScaleWidth      =   9555
      TabIndex        =   21
      Top             =   1440
      Visible         =   0   'False
      Width           =   9615
      Begin VB.PictureBox Shape1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":08CA
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   0
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   1
         Left            =   960
         MouseIcon       =   "frmMain.frx":2C9B
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   0
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   2
         Left            =   1920
         Picture         =   "frmMain.frx":2DED
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   0
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   3
         Left            =   2880
         MouseIcon       =   "frmMain.frx":51BE
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   0
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   4
         Left            =   3840
         Picture         =   "frmMain.frx":5310
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   0
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   5
         Left            =   4800
         MouseIcon       =   "frmMain.frx":76E1
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   0
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   6
         Left            =   5760
         Picture         =   "frmMain.frx":7833
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   0
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   7
         Left            =   6720
         MouseIcon       =   "frmMain.frx":9C04
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   0
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   8
         Left            =   7680
         Picture         =   "frmMain.frx":9D56
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   0
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   9
         Left            =   8640
         MouseIcon       =   "frmMain.frx":C127
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   0
         Width           =   975
         Begin VB.Timer Timer1 
            Interval        =   100
            Left            =   240
            Top             =   360
         End
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   10
         Left            =   0
         MouseIcon       =   "frmMain.frx":C279
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   960
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   11
         Left            =   960
         Picture         =   "frmMain.frx":C3CB
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   960
         Width           =   975
         Begin VB.Shape Shape9 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   4
            FillColor       =   &H00FFFFFF&
            Height          =   1095
            Left            =   13200
            Top             =   -240
            Width           =   14895
         End
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   12
         Left            =   1920
         MouseIcon       =   "frmMain.frx":E79C
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   960
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   13
         Left            =   2880
         Picture         =   "frmMain.frx":E8EE
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   960
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   14
         Left            =   3840
         MouseIcon       =   "frmMain.frx":10CBF
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   960
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   15
         Left            =   4800
         Picture         =   "frmMain.frx":10E11
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   960
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   16
         Left            =   5760
         MouseIcon       =   "frmMain.frx":131E2
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   960
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   17
         Left            =   6720
         Picture         =   "frmMain.frx":13334
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   960
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   18
         Left            =   7680
         MouseIcon       =   "frmMain.frx":15705
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   960
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   19
         Left            =   8640
         Picture         =   "frmMain.frx":15857
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   960
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   20
         Left            =   0
         Picture         =   "frmMain.frx":17C28
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   1920
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   21
         Left            =   960
         MouseIcon       =   "frmMain.frx":19FF9
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   1920
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   22
         Left            =   1920
         Picture         =   "frmMain.frx":1A14B
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   1920
         Width           =   975
         Begin VB.Shape Shape10 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   4
            FillColor       =   &H00FFFFFF&
            Height          =   1095
            Left            =   14040
            Top             =   720
            Width           =   14895
         End
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   23
         Left            =   2880
         MouseIcon       =   "frmMain.frx":1C51C
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   1920
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   24
         Left            =   3840
         Picture         =   "frmMain.frx":1C66E
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   1920
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   25
         Left            =   4800
         MouseIcon       =   "frmMain.frx":1EA3F
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   1920
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   26
         Left            =   5760
         Picture         =   "frmMain.frx":1EB91
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   1920
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   27
         Left            =   6720
         MouseIcon       =   "frmMain.frx":20F62
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   1920
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   28
         Left            =   7680
         Picture         =   "frmMain.frx":210B4
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   1920
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   29
         Left            =   8640
         MouseIcon       =   "frmMain.frx":23485
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   1920
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   30
         Left            =   0
         MouseIcon       =   "frmMain.frx":235D7
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   2880
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   31
         Left            =   960
         Picture         =   "frmMain.frx":23729
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   2880
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   32
         Left            =   1920
         MouseIcon       =   "frmMain.frx":25AFA
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   2880
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   33
         Left            =   2880
         Picture         =   "frmMain.frx":25C4C
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   2880
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   34
         Left            =   3840
         MouseIcon       =   "frmMain.frx":2801D
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   2880
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   35
         Left            =   4800
         Picture         =   "frmMain.frx":2816F
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   2880
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   36
         Left            =   5760
         MouseIcon       =   "frmMain.frx":2A540
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   2880
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   37
         Left            =   6720
         Picture         =   "frmMain.frx":2A692
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   2880
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   38
         Left            =   7680
         MouseIcon       =   "frmMain.frx":2CA63
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   2880
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   39
         Left            =   8640
         Picture         =   "frmMain.frx":2CBB5
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   2880
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   40
         Left            =   0
         Picture         =   "frmMain.frx":2EF86
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   3840
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   41
         Left            =   960
         MouseIcon       =   "frmMain.frx":31357
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   3840
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   42
         Left            =   1920
         Picture         =   "frmMain.frx":314A9
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   3840
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   43
         Left            =   2880
         MouseIcon       =   "frmMain.frx":3387A
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   3840
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   44
         Left            =   3840
         Picture         =   "frmMain.frx":339CC
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   3840
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   45
         Left            =   4800
         MouseIcon       =   "frmMain.frx":35D9D
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   3840
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   46
         Left            =   5760
         Picture         =   "frmMain.frx":35EEF
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   3840
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   47
         Left            =   6720
         MouseIcon       =   "frmMain.frx":382C0
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   3840
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   48
         Left            =   7680
         Picture         =   "frmMain.frx":38412
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   3840
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   49
         Left            =   8640
         MouseIcon       =   "frmMain.frx":3A7E3
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   3840
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   50
         Left            =   0
         MouseIcon       =   "frmMain.frx":3A935
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   4800
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   51
         Left            =   960
         Picture         =   "frmMain.frx":3AA87
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   4800
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   52
         Left            =   1920
         MouseIcon       =   "frmMain.frx":3CE58
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   4800
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   53
         Left            =   2880
         Picture         =   "frmMain.frx":3CFAA
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   4800
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   54
         Left            =   3840
         MouseIcon       =   "frmMain.frx":3F37B
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   4800
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   55
         Left            =   4800
         Picture         =   "frmMain.frx":3F4CD
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   4800
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   56
         Left            =   5760
         MouseIcon       =   "frmMain.frx":4189E
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   4800
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   57
         Left            =   6720
         Picture         =   "frmMain.frx":419F0
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   4800
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   58
         Left            =   7680
         MouseIcon       =   "frmMain.frx":43DC1
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   4800
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   59
         Left            =   8640
         Picture         =   "frmMain.frx":43F13
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   4800
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   60
         Left            =   0
         Picture         =   "frmMain.frx":462E4
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   5760
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   61
         Left            =   960
         MouseIcon       =   "frmMain.frx":486B5
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   5760
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   62
         Left            =   1920
         Picture         =   "frmMain.frx":48807
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   5760
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   63
         Left            =   2880
         MouseIcon       =   "frmMain.frx":4ABD8
         MousePointer    =   99  'Custom
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   5760
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   64
         Left            =   3840
         Picture         =   "frmMain.frx":4AD2A
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   5760
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   65
         Left            =   4800
         MouseIcon       =   "frmMain.frx":4D0FB
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   5760
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   66
         Left            =   5760
         Picture         =   "frmMain.frx":4D24D
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   5760
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   67
         Left            =   6720
         MouseIcon       =   "frmMain.frx":4F61E
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   5760
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   68
         Left            =   7680
         Picture         =   "frmMain.frx":4F770
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   5760
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   69
         Left            =   8640
         MouseIcon       =   "frmMain.frx":51B41
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   5760
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   70
         Left            =   0
         MouseIcon       =   "frmMain.frx":51C93
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   6720
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   71
         Left            =   960
         Picture         =   "frmMain.frx":51DE5
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   6720
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   72
         Left            =   1920
         MouseIcon       =   "frmMain.frx":541B6
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   6720
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   73
         Left            =   2880
         Picture         =   "frmMain.frx":54308
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   6720
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   74
         Left            =   3840
         MouseIcon       =   "frmMain.frx":566D9
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   6720
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   75
         Left            =   4800
         Picture         =   "frmMain.frx":5682B
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   6720
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   76
         Left            =   5760
         MouseIcon       =   "frmMain.frx":58BFC
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   6720
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   77
         Left            =   6720
         Picture         =   "frmMain.frx":58D4E
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   6720
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   78
         Left            =   7680
         MouseIcon       =   "frmMain.frx":5B11F
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   6720
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   79
         Left            =   8640
         Picture         =   "frmMain.frx":5B271
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   6720
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   80
         Left            =   0
         Picture         =   "frmMain.frx":5D642
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   7680
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   81
         Left            =   960
         MouseIcon       =   "frmMain.frx":5FA13
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   7680
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   82
         Left            =   1920
         Picture         =   "frmMain.frx":5FB65
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   7680
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   83
         Left            =   2880
         MouseIcon       =   "frmMain.frx":61F36
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   7680
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   84
         Left            =   3840
         Picture         =   "frmMain.frx":62088
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   7680
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   85
         Left            =   4800
         MouseIcon       =   "frmMain.frx":64459
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   7680
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   86
         Left            =   5760
         Picture         =   "frmMain.frx":645AB
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   7680
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   87
         Left            =   6720
         MouseIcon       =   "frmMain.frx":6697C
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   7680
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   88
         Left            =   7680
         Picture         =   "frmMain.frx":66ACE
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   7680
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   89
         Left            =   8640
         MouseIcon       =   "frmMain.frx":68E9F
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   7680
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   90
         Left            =   0
         MouseIcon       =   "frmMain.frx":68FF1
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   8640
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   91
         Left            =   960
         Picture         =   "frmMain.frx":69143
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   8640
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   92
         Left            =   1920
         MouseIcon       =   "frmMain.frx":6B514
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   8640
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   93
         Left            =   2880
         Picture         =   "frmMain.frx":6B666
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   8640
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   94
         Left            =   3840
         MouseIcon       =   "frmMain.frx":6DA37
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   8640
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   95
         Left            =   4800
         Picture         =   "frmMain.frx":6DB89
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   8640
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   96
         Left            =   5760
         MouseIcon       =   "frmMain.frx":6FF5A
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   8640
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   97
         Left            =   6720
         Picture         =   "frmMain.frx":700AC
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   8640
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   98
         Left            =   7680
         MouseIcon       =   "frmMain.frx":7247D
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   8640
         Width           =   975
      End
      Begin VB.PictureBox Shape1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   99
         Left            =   8640
         Picture         =   "frmMain.frx":725CF
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   8640
         Width           =   975
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   495
         Left            =   7920
         TabIndex        =   143
         Top             =   6840
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
         Max             =   5
      End
      Begin VB.Frame Frame3 
         Caption         =   "Advanced Options"
         Height          =   2175
         Left            =   7920
         TabIndex        =   122
         Top             =   2760
         Visible         =   0   'False
         Width           =   1695
         Begin VB.TextBox txtDepth 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   240
            TabIndex        =   125
            TabStop         =   0   'False
            ToolTipText     =   "Click to change"
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtThreshold 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   120
            TabIndex        =   124
            TabStop         =   0   'False
            ToolTipText     =   "Click to change"
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtMaxTime 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   240
            TabIndex        =   123
            TabStop         =   0   'False
            ToolTipText     =   "Click to change"
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Labels 
            Alignment       =   2  'Center
            Caption         =   "Max Ply Depth"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   128
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Labels 
            Alignment       =   2  'Center
            Caption         =   "Pruning Threshold"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   127
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Labels 
            Alignment       =   2  'Center
            Caption         =   "Max Think Time"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   126
            Top             =   1440
            Width           =   1455
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Game Info"
         Height          =   2655
         Left            =   7920
         TabIndex        =   129
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
         Begin VB.Label lblP1Time 
            Alignment       =   2  'Center
            Caption         =   "0 Min 0 Sec"
            Height          =   255
            Left            =   120
            TabIndex        =   137
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lblP2Time 
            Alignment       =   2  'Center
            Caption         =   "0 Min 0 Sec"
            Height          =   255
            Left            =   120
            TabIndex        =   136
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblTotalTime 
            Alignment       =   2  'Center
            Caption         =   "0 Min 0 Sec"
            Height          =   255
            Left            =   120
            TabIndex        =   135
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label lblTurns 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Left            =   120
            TabIndex        =   134
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Labels 
            Alignment       =   2  'Center
            Caption         =   "P1 Time"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   133
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Labels 
            Alignment       =   2  'Center
            Caption         =   "P2 Time"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   132
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Labels 
            Alignment       =   2  'Center
            Caption         =   "Total Time"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   131
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Labels 
            Alignment       =   2  'Center
            Caption         =   "Total Turns"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   130
            Top             =   2040
            Width           =   1455
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Statistics"
         Height          =   1455
         Left            =   7920
         TabIndex        =   138
         Top             =   5040
         Visible         =   0   'False
         Width           =   1695
         Begin VB.Label lblPlyDepth 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Left            =   120
            TabIndex        =   142
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lblMMatrixSize 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Left            =   120
            TabIndex        =   141
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Labels 
            Alignment       =   2  'Center
            Caption         =   "Current Depth"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   140
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Labels 
            Alignment       =   2  'Center
            Caption         =   "Current Moves"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   139
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Label Labels 
         Alignment       =   2  'Center
         Caption         =   "Move Speed"
         Height          =   255
         Index           =   4
         Left            =   7920
         TabIndex        =   144
         Top             =   6600
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdDebug 
      Caption         =   "Auto &Debug"
      Height          =   375
      Left            =   600
      TabIndex        =   20
      Top             =   10560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Start"
      Height          =   375
      Left            =   600
      TabIndex        =   19
      Top             =   10560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdName 
      Caption         =   "&Name Players"
      Height          =   375
      Left            =   600
      TabIndex        =   18
      Top             =   10560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   375
      Left            =   600
      TabIndex        =   17
      Top             =   10560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2640
      Top             =   10200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   65
      ImageHeight     =   65
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":749A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":77BB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7ADD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7DFE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":81200
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3240
      Top             =   10320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Board"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   10560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Load Board"
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   10560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   15
      Top             =   11190
      Visible         =   0   'False
      Width           =   15330
      _ExtentX        =   27040
      _ExtentY        =   529
      Style           =   1
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gameplay Type"
      Height          =   975
      Left            =   4080
      TabIndex        =   10
      Top             =   6840
      Visible         =   0   'False
      Width           =   1575
      Begin VB.OptionButton Option1 
         Caption         =   "1 Player"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "2 Player"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "Skip &Go"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   10560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdReverse 
      Caption         =   "&Reverse Board"
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   10560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   10560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Options"
      Height          =   1695
      Left            =   4080
      TabIndex        =   14
      Top             =   9120
      Visible         =   0   'False
      Width           =   1575
      Begin VB.CheckBox CheckForce 
         Caption         =   "Force Taking"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox CheckABP 
         Caption         =   "ABP Mode"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox CheckCheat 
         Caption         =   "Cheat"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox CheckAutoSwitch 
         Caption         =   "Auto Switch"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Points"
      Height          =   975
      Left            =   4080
      TabIndex        =   11
      Top             =   7920
      Visible         =   0   'False
      Width           =   1575
      Begin VB.Label lblP2Points 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblP1Points 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Image Image11 
      Height          =   9615
      Left            =   3120
      Picture         =   "frmMain.frx":84418
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   9615
   End
   Begin VB.Image Image10 
      Appearance      =   0  'Flat
      Height          =   1095
      Left            =   13560
      Picture         =   "frmMain.frx":A0EF6
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Image9 
      Appearance      =   0  'Flat
      Height          =   1095
      Left            =   720
      Picture         =   "frmMain.frx":AE99B
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      X1              =   856
      X2              =   872
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      X1              =   856
      X2              =   872
      Y1              =   440
      Y2              =   440
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      X1              =   856
      X2              =   872
      Y1              =   264
      Y2              =   264
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   9615
      Left            =   13080
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      X1              =   160
      X2              =   200
      Y1              =   680
      Y2              =   680
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      X1              =   160
      X2              =   200
      Y1              =   544
      Y2              =   544
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      X1              =   160
      X2              =   200
      Y1              =   416
      Y2              =   416
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      X1              =   160
      X2              =   200
      Y1              =   288
      Y2              =   288
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      X1              =   160
      X2              =   200
      Y1              =   160
      Y2              =   160
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   9855
      Left            =   3000
      Top             =   1320
      Width           =   9855
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PINOY DAMA"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   1440
      TabIndex        =   154
      Top             =   0
      Width           =   13815
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   1575
      Left            =   360
      Top             =   9360
      Width           =   1935
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   1575
      Left            =   360
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   1575
      Left            =   360
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   1575
      Left            =   360
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   4
      FillColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   360
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   480
      MouseIcon       =   "frmMain.frx":BC440
      MousePointer    =   99  'Custom
      TabIndex        =   149
      Top             =   9960
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  HELP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   600
      MouseIcon       =   "frmMain.frx":BC592
      MousePointer    =   99  'Custom
      TabIndex        =   148
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  SAVE GAME"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1215
      Left            =   720
      MouseIcon       =   "frmMain.frx":BC9D4
      MousePointer    =   99  'Custom
      TabIndex        =   147
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  OPEN GAME"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   720
      MouseIcon       =   "frmMain.frx":BCB26
      MousePointer    =   99  'Custom
      TabIndex        =   146
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  NEW GAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1455
      Left            =   480
      MouseIcon       =   "frmMain.frx":BCC78
      MousePointer    =   99  'Custom
      TabIndex        =   145
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Image Image4 
      Height          =   1590
      Left            =   360
      MouseIcon       =   "frmMain.frx":BCDCA
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":BD20C
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   1905
   End
   Begin VB.Image Image3 
      Height          =   1590
      Left            =   360
      MouseIcon       =   "frmMain.frx":CD3EB
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":CD82D
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1905
   End
   Begin VB.Image Image2 
      Height          =   1590
      Left            =   360
      MouseIcon       =   "frmMain.frx":DDA0C
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":DDE4E
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   1905
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   360
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":EE02D
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1905
   End
   Begin VB.Image Image5 
      Height          =   1590
      Left            =   360
      MouseIcon       =   "frmMain.frx":FE20C
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":FE64E
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   1905
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   9615
      Left            =   240
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PINOY DAMA"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   1320
      TabIndex        =   153
      Top             =   0
      Width           =   13815
   End
   Begin VB.Image Image7 
      Height          =   11415
      Left            =   0
      Picture         =   "frmMain.frx":10E82D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15375
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      X1              =   160
      X2              =   200
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Image Image8 
      Height          =   11415
      Left            =   0
      Picture         =   "frmMain.frx":133459
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15375
   End
   Begin VB.Image Image6 
      Height          =   11415
      Left            =   0
      Picture         =   "frmMain.frx":158085
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15375
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SndPlayed As Boolean
Option Explicit


Private Sub CheckABP_Click()
  ABPMode = CheckABP
  General.SaveSettings
End Sub

Private Sub CheckABP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to toggle Alpha Beta Pruning mode"
End Sub

Private Sub CheckAutoSwitch_Click()
  AutoSwitch = CheckAutoSwitch
  If AutoSwitch = 1 Then cmdReverse.Enabled = False Else cmdReverse.Enabled = True
  Select Case Turn
    Case 1
      If Reversed = True Then Reversed = False: RefreshBoard Currentpieces
    Case 2
      If Reversed = False Then Reversed = True: RefreshBoard Currentpieces
  End Select
  
  General.SaveSettings
End Sub

Private Sub CheckAutoSwitch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to automatically switch the board orientation to that of the current players' view (only in 2 player mode)"
End Sub

Private Sub CheckCheat_Click()
  CheatSwitch = CheckCheat
  General.SaveSettings
End Sub

Private Sub CheckCheat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Used as debug, allow you to move any piece anywhere, moving a piece onto another deletes the original piece"
End Sub

Private Sub CheckForce_Click()
  Select Case CheckForce
    Case 0
      ForceMove = False
    Case 1
      ForceMove = True
  End Select
  General.SaveSettings
End Sub

Private Sub CheckForce_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Toggles Force Taking Mode, this rule states that all players must take the maximum amount of pieces possible for their move"
End Sub

Private Sub cmdBack_Click()
Dim Lng1 As Long
For Lng1 = 1 To 40
  Currentpieces(Lng1) = Lastpieces(Lng1)
Next
  RefreshBoard Currentpieces
End Sub

Private Sub cmdBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to revert back to board state before last move"
End Sub

Private Sub cmdDebug_Click()
  AutoDebug = Not AutoDebug
  If AutoDebug = True Then Call AIMove(Turn)
End Sub

Private Sub cmdDebug_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to toggle Automatic Debug Mode (the computer plays itself until an error occurs)"
End Sub

Private Sub cmdExit_Click()
  Unload Me
  End
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to exit game"
End Sub

Private Sub cmdName_Click()
Dim Response As String

Response = InputBox("Enter Player 1 Name", "Player 1", Names(1))
If Response <> "" Then Names(1) = Response Else Exit Sub

Response = InputBox("Enter Player 2 Name", "Player 2", Names(2))
If Response <> "" Then Names(2) = Response Else Exit Sub
Label2.Enabled = True
Label3.Enabled = True
board1.Visible = True
turn1.Visible = True
General.SaveSettings
General.RefreshDisplay

End Sub

Private Sub cmdName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to name players"
End Sub

Private Sub cmdOpen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to load a saved game"
End Sub

Private Sub cmdReset_Click()
  Call General.ResetGame
End Sub

Private Sub cmdReset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Starts a new game"
End Sub

Private Sub cmdReverse_Click()
  Reversed = Not Reversed
  Call RefreshBoard(Currentpieces)
End Sub

Private Sub cmdReverse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Reverses board orientation"
End Sub

Private Sub cmdOpen_Click()
Dim File As String, Response As Long, Buffer As String, Gametype As Long
  If GameStarted Then
    Response = MsgBox("Exit current game?", vbExclamation + vbYesNo, "PINOY DAMA")
    If Response = vbNo Then Exit Sub
  End If
  
  CD1.DialogTitle = "Load Board"
  CD1.Filter = "Board File (*.dma)|*.dma"
  CD1.Flags = cdlOFNHideReadOnly
  CD1.InitDir = CurDir
  CD1.FileName = ""
  CD1.ShowOpen
  File = CD1.FileName
  If File = "" Then Exit Sub
  
  If Dir(File, vbHidden Or vbSystem Or vbNormal Or vbReadOnly) = "" Then
    MsgBox CD1.FileTitle & " does not exist", vbExclamation
    Exit Sub
  End If
  
  Open File For Binary Access Read As #1
    Get #1, , Currentpieces
    Get #1, , Turn
    Get #1, , Gametype
  Close #1
  
  GameStarted = True
  
  If (Gametype) Then Option1 = True Else Option2 = True
  
  General.RefreshDisplay
  General.RefreshBoard Currentpieces
  
End Sub

Private Sub cmdSave_Click()
Dim File As String, Response As Long, Buffer As String

  CD1.DialogTitle = "Save Board"
  CD1.Filter = "Board File (*.dma)|*.dma"
  CD1.Flags = cdlOFNHideReadOnly
  CD1.InitDir = CurDir
  CD1.FileName = ""
  CD1.ShowSave
  File = CD1.FileName
  If File = "" Then Exit Sub
  If Dir(File, vbHidden Or vbSystem Or vbNormal Or vbReadOnly) <> "" Then
    Response = MsgBox(CD1.FileTitle & " already exists, overwrite?", vbExclamation + vbYesNo)
    If Response = vbNo Then Exit Sub
  End If
  
  Open File For Binary As #1
    Put #1, , Currentpieces
    Put #1, , Turn
    Put #1, , CLng(Option1)
  Close #1
  
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to save the current game"
End Sub

Private Sub cmdSwitch_Click()
  Select Case Turn
    Case 1
      Turn = 2
      If Option1 = True Then Call AIMove(2)
    Case 2
      Turn = 1
  End Select
  RefreshDisplay
End Sub

Private Sub cmdSwitch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to skip the current players go"
End Sub

Private Sub Form_Load()
 
 
 CheckForce.Value = 1
 cmdReverse.Enabled = True
 PlayerType = 2
 General.SaveSettings
Dim Lng1 As Long, Lng2 As Long
  
  
  IndexMoves(1, 1) = -9
  IndexMoves(2, 1) = 11
  IndexMoves(3, 1) = 9
  IndexMoves(4, 1) = -11
  IndexMoves(5, 1) = -18
  IndexMoves(6, 1) = 22
  IndexMoves(7, 1) = 18
  IndexMoves(8, 1) = -22
  
  XYMoves(1, 1).X = 1
  XYMoves(2, 1).X = 1
  XYMoves(3, 1).X = -1
  XYMoves(4, 1).X = -1
  XYMoves(5, 1).X = 2
  XYMoves(6, 1).X = 2
  XYMoves(7, 1).X = -2
  XYMoves(8, 1).X = -2
  
  XYMoves(1, 1).Y = -1
  XYMoves(2, 1).Y = 1
  XYMoves(3, 1).Y = 1
  XYMoves(4, 1).Y = -1
  XYMoves(5, 1).Y = -2
  XYMoves(6, 1).Y = 2
  XYMoves(7, 1).Y = 2
  XYMoves(8, 1).Y = -2
  
  For Lng2 = 1 To 8
    For Lng1 = 1 To 4
      IndexMoves(Lng1, Lng2) = IndexMoves(Lng1, 1) * Lng2
      XYMoves(Lng1, Lng2).X = XYMoves(Lng1, 1).X * Lng2
      XYMoves(Lng1, Lng2).Y = XYMoves(Lng1, 1).Y * Lng2
    Next Lng1
    For Lng1 = 5 To 8
      IndexMoves(Lng1, Lng2) = IndexMoves(Lng1 - 4, 1) * (Lng2 + 1)
      XYMoves(Lng1, Lng2).X = XYMoves(Lng1 - 4, 1).X * (Lng2 + 1)
      XYMoves(Lng1, Lng2).Y = XYMoves(Lng1 - 4, 1).Y * (Lng2 + 1)
    Next
  Next Lng2
  
  ReDim SMatrix(0 To 0)
  ReDim MoveMatrix(0 To 0)

  General.GetSettings
  
  If Names(1) = "" Then
    Names(1) = "Player 1"
    Names(2) = "Nemesis"
    MaxDepth = 5
    txtDepth = "5"
    PruneThreshold = 40
    txtThreshold = "25%"
    PlayerType = 1
    ForceMove = True
    General.SaveSettings
  End If
  
  If PlayerType = 1 Then Option1 = True Else Option2 = True
  If ForceMove = True Then CheckForce = 1 Else CheckForce = 0
  CheckAutoSwitch = AutoSwitch
  CheckCheat = CheatSwitch
  CheckABP = ABPMode
  txtDepth = MaxDepth
  txtThreshold = PruneThreshold
  txtMaxTime = MaxThoughtTime & " Sec"
  Select Case MoveSpeed
    Case 50
      Slider1 = 5
    Case 100
      Slider1 = 4
    Case 200
      Slider1 = 3
    Case 300
      Slider1 = 2
    Case 400
      Slider1 = 1
    Case 500
      Slider1 = 0
  End Select
  
  General.ResetGame
  General.RefreshDisplay
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  StatusBar1.SimpleText = ""
End Sub

Private Sub Label1_Click()
    If SndPlayed = False Then
        PlaySnd (TheButton) 'plays the sound
        SndPlayed = True
    End If

Call cmdName_Click
End Sub


Private Sub Label2_Click()
    If SndPlayed = False Then
        PlaySnd (TheButton) 'plays the sound
        SndPlayed = True
    End If

Call cmdOpen_Click
End Sub

Private Sub Label3_Click()
    If SndPlayed = False Then
        PlaySnd (TheButton) 'plays the sound
        SndPlayed = True
    End If

Call cmdSave_Click
End Sub

Private Sub Label4_Click()
    If SndPlayed = False Then
        PlaySnd (TheButton) 'plays the sound
        SndPlayed = True
    End If

help.Show
End Sub

Private Sub Label5_Click()
    If SndPlayed = False Then
        PlaySnd (TheButton) 'plays the sound
        SndPlayed = True
    End If

End
End Sub

Private Sub lblMMatrixSize_Click()
  StatusBar1.SimpleText = "Displays the amount of moves the AI engine has generated so far"
End Sub

Private Sub lblP1Time_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Displays time taken for " & Names(1) & " this game"
End Sub

Private Sub lblP2Time_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Displays time taken for " & Names(2) & " this game"
End Sub

Private Sub lblPlyDepth_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Displays current depth that the AI engine is thinking at"
End Sub

Private Sub lblTotalTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Displays total time taken for this game"
End Sub


Private Sub lblTurns_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Displays number of turns this game"
End Sub

Private Sub Option1_Click()
  If Turn = 2 Then Call AIMove(2)
  cmdReverse.Enabled = False
  PlayerType = 1
  General.SaveSettings
End Sub

Private Sub Option2_Click()
  cmdReverse.Enabled = True
  PlayerType = 2
  General.SaveSettings
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Shape1_Click(Index As Integer)
Static Selected As Boolean, SelectedPiece As SelectedSquare, RealIndex As Long
Dim Direction As Long, MoveLength As Long, MaxMoveLength As Long

RealIndex = General.IndexTranslation(Val(Index))

If Option1 = True And Turn = 2 Then Exit Sub

If Selected = True Then
  If SelectedPiece.Index = RealIndex And (P1MultiMode = False Or ForceMove = False) Then GoTo Deselect
  GameStarted = True
  If MovePiece(Val(SelectedPiece.Index), RealIndex, 0, False) Then Selected = False
  If P1MultiMode Then SelectedPiece = CheckSquare(Currentpieces, , , Val(Currentpieces(SelectedPiece.Piece).Index))
Else
  SelectedPiece = CheckSquare(Currentpieces, , , Val(RealIndex))
  
  If ForceMove = True Then
    If SelectedPiece.IsPiece = False Then Exit Sub
    If SelectedPiece.Double = True Then MaxMoveLength = 8 Else MaxMoveLength = 1
    For Direction = 1 To 4
      For MoveLength = 1 To MaxMoveLength
        If SelectedPiece.Index + IndexMoves(Direction, MoveLength) > 99 Or SelectedPiece.Index + IndexMoves(Direction, MoveLength) < 0 Then Exit For
        If MovePiece(SelectedPiece.Index, SelectedPiece.Index + IndexMoves(Direction, MoveLength), 0, False, True) = True Then
          If SelectedPiece.IsPiece And (Turn = SelectedPiece.Player Or CheatSwitch = 1) Then
            Shape1(Index).Picture = ImageList1.ListImages(5).Picture
            Selected = True
          End If
        End If
      Next
    Next
  Else
    If SelectedPiece.IsPiece And (Turn = SelectedPiece.Player Or CheatSwitch = 1) Then
      Shape1(Index).Picture = ImageList1.ListImages(5).Picture
      Selected = True
    End If
  End If
End If

Exit Sub

Deselect:

If P1MultiMode = True And ((SelectedPiece.Player = 1 And SelectedPiece.Y = 1) Or (SelectedPiece.Player = 2 And SelectedPiece.Y = 8)) Then Currentpieces(SelectedPiece.Piece).Double = True: SelectedPiece.Double = True

If SelectedPiece.Player = 1 Then
  If SelectedPiece.Double Then
    Shape1(Index).Picture = ImageList1.ListImages(2).Picture
  Else
    Shape1(Index).Picture = ImageList1.ListImages(1).Picture
  End If
Else
  If SelectedPiece.Double Then
    Shape1(Index).Picture = ImageList1.ListImages(4).Picture
  Else
    Shape1(Index).Picture = ImageList1.ListImages(3).Picture
  End If
End If

Selected = False

If P1MultiMode Then
  P1MultiMode = False
  If Turn = 1 Then
    Turn = 2
    If CheckWin(Currentpieces) = True Then
      RefreshDisplay
    Else
      RefreshDisplay
    End If
    DoEvents
    If frmmain.Option1 And CheatSwitch = 0 Then Call AIMove(CByte(Turn))
    If frmmain.Option2 And frmmain.CheckAutoSwitch Then Sleep 500: Reversed = True: RefreshBoard Currentpieces  ' Switch board
  Else
    Turn = 1
    If CheckWin(Currentpieces) = True Then
      RefreshDisplay
    Else
      RefreshDisplay
    End If
    DoEvents
    If frmmain.Option2 And frmmain.CheckAutoSwitch Then Sleep 500: Reversed = False: RefreshBoard Currentpieces  ' Switch board
  End If
End If

End Sub

Private Sub Shape1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim X1 As Long, Y1 As Long, TempSquare As SelectedSquare, Msg As String
  Index = IndexTranslation(CLng(Index))
  XYConvert CLng(Index), X1, Y1
  TempSquare = CheckSquare(Currentpieces, , , CLng(Index))
  
  If TempSquare.IsPiece Then
    If TempSquare.Player = 1 Then
     Msg = "     " & Names(1)
    Else
     Msg = "     " & Names(2)
    End If
    If TempSquare.Double Then Msg = Msg & " Double Piece" Else Msg = Msg & " Single Piece"
  End If
  
  StatusBar1.SimpleText = "Index = " & Index & "     X = " & X1 & "     Y = " & Y1 & Msg
End Sub

Private Sub Slider1_Scroll()
  Select Case Slider1
      Case 0
        Slider1.Text = "Slowest"
        MoveSpeed = 500
      Case 1
        Slider1.Text = "Slow"
        MoveSpeed = 400
      Case 2
        Slider1.Text = "Normal"
        MoveSpeed = 300
      Case 3
        Slider1.Text = "Fast"
        MoveSpeed = 200
      Case 4
        Slider1.Text = "Fastest"
        MoveSpeed = 100
      Case 5
        Slider1.Text = "Insane"
        MoveSpeed = 50
    End Select
    General.SaveSettings
End Sub

Private Sub Timer1_Timer()

  If GameStarted = False Then Exit Sub
  Select Case Turn
    Case 1
      VP1Time.Seconds = VP1Time.Seconds + 0.1
      If VP1Time.Seconds >= 60 Then VP1Time.Minutes = VP1Time.Minutes + Int(VP1Time.Seconds / 60): VP1Time.Seconds = VP2Time.Seconds - (Int(VP2Time.Seconds / 60) * 60)
      If InStr(1, CStr(Round(VP1Time.Seconds, 1)), ".", vbBinaryCompare) = 0 Then
        lblP1Time = VP1Time.Minutes & " Min " & Round(VP1Time.Seconds, 1) & ".0 Sec"
      Else
        lblP1Time = VP1Time.Minutes & " Min " & Round(VP1Time.Seconds, 1) & " Sec"
      End If
    Case 2
      If Option1 = True Then Exit Sub
      VP2Time.Seconds = VP2Time.Seconds + 0.1
      If VP2Time.Seconds >= 60 Then VP2Time.Minutes = VP2Time.Minutes + Int(VP2Time.Seconds / 60): VP2Time.Seconds = VP2Time.Seconds - (Int(VP2Time.Seconds / 60) * 60)
      If InStr(1, CStr(Round(VP2Time.Seconds, 1)), ".", vbBinaryCompare) = 0 Then
        lblP2Time = VP2Time.Minutes & " Min " & Round(VP2Time.Seconds, 1) & ".0 Sec"
      Else
        lblP2Time = VP2Time.Minutes & " Min " & Round(VP2Time.Seconds, 1) & " Sec"
      End If
  End Select
  
  TotalTime.Seconds = VP2Time.Seconds + VP1Time.Seconds
  TotalTime.Minutes = VP2Time.Minutes + VP2Time.Minutes
  If TotalTime.Seconds >= 60 Then TotalTime.Minutes = TotalTime.Minutes + Int(TotalTime.Seconds / 60): TotalTime.Seconds = TotalTime.Seconds - (Int(TotalTime.Seconds / 60) * 60)
  If InStr(1, CStr(Round(TotalTime.Seconds, 1)), ".", vbBinaryCompare) = 0 Then
    lblTotalTime = TotalTime.Minutes & " Min " & Round(TotalTime.Seconds, 1) & ".0 Sec"
  Else
    lblTotalTime = TotalTime.Minutes & " Min " & Round(TotalTime.Seconds, 1) & " Sec"
  End If
  
End Sub

Private Sub txtDepth_Click()
  txtDepth.Alignment = 0
  txtDepth.BackColor = &H8000000E
  txtDepth.BorderStyle = 1
  txtDepth.SelStart = 0
  txtDepth.SelLength = Len(txtDepth)
End Sub

Private Sub txtDepth_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call txtDepth_LostFocus: cmdReset.SetFocus
End Sub

Private Sub txtDepth_LostFocus()
  txtDepth.Alignment = 2
  txtDepth.BackColor = &H8000000F
  txtDepth.BorderStyle = 0
  If IsNumeric(txtDepth) Then
    If Val(txtDepth) > 1 Then
      MaxDepth = txtDepth
      txtDepth = MaxDepth
      General.SaveSettings
    Else
      MsgBox "The maximum ply depth must be greater than 1", vbExclamation
      txtDepth = MaxDepth
      txtDepth.SelStart = 0
      txtDepth.SelLength = Len(txtDepth)
    End If
  Else
    MsgBox "The Maximum Ply Depth must be numeric!", vbExclamation
    txtDepth = MaxDepth
    txtDepth.SelStart = 0
    txtDepth.SelLength = Len(txtDepth)
  End If
End Sub

Private Sub txtDepth_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to change maximum ply depth"
End Sub

Private Sub txtThreshold_Click()
  txtThreshold.Alignment = 0
  txtThreshold.BackColor = &H8000000E
  txtThreshold.BorderStyle = 1
  txtThreshold.SelStart = 0
  txtThreshold.SelLength = Len(txtThreshold)
End Sub

Private Sub txtThreshold_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call txtThreshold_LostFocus: cmdReset.SetFocus
End Sub

Private Sub txtThreshold_LostFocus()
  txtThreshold.Alignment = 2
  txtThreshold.BackColor = &H8000000F
  txtThreshold.BorderStyle = 0
  If IsNumeric(txtThreshold) Then
    PruneThreshold = txtThreshold
    General.SaveSettings
  Else
    MsgBox "The Array Upper Bound Error Margin must be numeric", vbExclamation
    txtThreshold = PruneThreshold
    txtThreshold.SelStart = 0
    txtThreshold.SelLength = Len(txtThreshold)
  End If
End Sub

Private Sub txtthreshold_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to change Alpha Beta Pruning threshold, score at which it prunes (100 = even score)"
End Sub

Private Sub txtMaxTime_Click()
  txtMaxTime = Left(txtMaxTime, Len(txtMaxTime) - 4)
  txtMaxTime.Alignment = 0
  txtMaxTime.BackColor = &H8000000E
  txtMaxTime.BorderStyle = 1
  txtMaxTime.SelStart = 0
  txtMaxTime.SelLength = Len(txtMaxTime)
End Sub

Private Sub txtMaxTime_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then cmdReset.SetFocus
End Sub

Private Sub txtMaxTime_LostFocus()
  txtMaxTime.Alignment = 2
  txtMaxTime.BackColor = &H8000000F
  txtMaxTime.BorderStyle = 0
  If IsNumeric(txtMaxTime) Then
    MaxThoughtTime = txtMaxTime
    General.SaveSettings
  Else
    MsgBox "The Array Upper Bound Error Margin must be numeric", vbExclamation
    txtMaxTime = PruneThreshold
    txtMaxTime.SelStart = 0
    txtMaxTime.SelLength = Len(txtMaxTime)
  End If
  txtMaxTime = txtMaxTime & " Sec"
End Sub

Private Sub txtMaxTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to set maximum tinking time for the computer player"
End Sub
