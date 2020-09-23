VERSION 5.00
Begin VB.Form about1 
   BackColor       =   &H80000007&
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8025
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image4 
      Height          =   495
      Left            =   7320
      Picture         =   "about1.frx":0000
      Stretch         =   -1  'True
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2004 nerd_boy17@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1020
      TabIndex        =   5
      Top             =   1200
      Width           =   3465
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   7440
      MouseIcon       =   "about1.frx":DAA5
      MousePointer    =   99  'Custom
      Picture         =   "about1.frx":DBF7
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BSCS 4A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   240
      Left            =   5625
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ronaldo C. Jornales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   240
      Left            =   5025
      TabIndex        =   3
      Top             =   600
      Width           =   2145
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jerry M. Jose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   240
      Left            =   5385
      TabIndex        =   2
      Top             =   840
      Width           =   1425
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROGRAMMERS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   240
      Left            =   4050
      TabIndex        =   1
      Top             =   240
      Width           =   1875
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PINOY DAMA"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   570
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   3045
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   240
      Picture         =   "about1.frx":E681
      Stretch         =   -1  'True
      Top             =   360
      Width           =   840
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   0
      Picture         =   "about1.frx":EF4B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "about1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image3_Click()
    If SndPlayed = False Then
        PlaySnd (TheButton) 'plays the sound
        SndPlayed = True
    End If

Unload Me
End Sub
