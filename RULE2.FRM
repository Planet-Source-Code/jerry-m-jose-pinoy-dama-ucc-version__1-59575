VERSION 5.00
Begin VB.Form rule2 
   BackColor       =   &H00000000&
   ClientHeight    =   5535
   ClientLeft      =   1260
   ClientTop       =   2820
   ClientWidth     =   8175
   ControlBox      =   0   'False
   Icon            =   "RULE2.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   5535
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image4 
      Height          =   495
      Left            =   7320
      Picture         =   "RULE2.frx":08CA
      Stretch         =   -1  'True
      Top             =   240
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   6000
      Picture         =   "RULE2.frx":1354
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rules of DAMA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"RULE2.frx":1796
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   2415
      Left            =   1200
      TabIndex        =   1
      Top             =   1680
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   720
      Shape           =   3  'Circle
      Top             =   1680
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   720
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   120
      X2              =   135
      Y1              =   3240
      Y2              =   5415
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   120
      X2              =   3840
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   6840
      Picture         =   "RULE2.frx":1880
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   4320
      Picture         =   "RULE2.frx":1CC2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "When a man ends its turn on the far row of the board it is crowned and becomes a King."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   975
      Left            =   1200
      TabIndex        =   0
      Top             =   3960
      Width           =   5175
   End
End
Attribute VB_Name = "rule2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MakeTransparent Me.hWnd, 200
End Sub

Private Sub Image2_Click()
    If SndPlayed = False Then
        PlaySnd (TheButton) 'plays the sound
        SndPlayed = True
    End If

rule1.Show
Me.Hide
End Sub

Private Sub Image3_Click()
    If SndPlayed = False Then
        PlaySnd (TheButton) 'plays the sound
        SndPlayed = True
    End If
rule3.Show
Me.Hide

End Sub

Private Sub Image4_Click()
    If SndPlayed = False Then
        PlaySnd (TheButton) 'plays the sound
        SndPlayed = True
    End If
Me.Hide
End Sub
