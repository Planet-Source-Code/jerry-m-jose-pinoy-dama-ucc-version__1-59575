VERSION 5.00
Begin VB.Form help 
   BackColor       =   &H80000007&
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10635
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "help.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6870
   ScaleWidth      =   10635
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image5 
      Height          =   1095
      Left            =   9120
      Picture         =   "help.frx":08CA
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   9840
      Picture         =   "help.frx":E36F
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABOUT AUTHORS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   360
      MouseIcon       =   "help.frx":EDF9
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RULES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   360
      MouseIcon       =   "help.frx":EF4B
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Image Image4 
      Height          =   1335
      Left            =   360
      Picture         =   "help.frx":F09D
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   1335
      Left            =   360
      Picture         =   "help.frx":1F27C
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   6825
      Left            =   0
      Picture         =   "help.frx":2F45B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10650
   End
End
Attribute VB_Name = "help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
    If SndPlayed = False Then
        PlaySnd (TheButton) 'plays the sound
        SndPlayed = True
    End If

Unload Me
End Sub

Private Sub Label1_Click()
    If SndPlayed = False Then
        PlaySnd (TheButton) 'plays the sound
        SndPlayed = True
    End If

rule1.Show
End Sub

Private Sub Label2_Click()
    If SndPlayed = False Then
        PlaySnd (TheButton) 'plays the sound
        SndPlayed = True
    End If

about1.Show
End Sub
