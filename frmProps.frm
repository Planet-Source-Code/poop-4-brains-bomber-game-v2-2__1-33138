VERSION 5.00
Begin VB.Form frmProps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Options"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2955
   Icon            =   "frmProps.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   2955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtSpeed 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Top Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed (1 to 5): "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
If Pause > 0 Then Pause = 1 'set the pause to 1 so if the game was running, it will resume
Unload Me 'unload this form
End Sub

Private Sub cmdReset_Click()
'reset the top score by saving over it with defaults
SaveSetting "BOMBER2", "TOPSCORE", "SCORE", "0"
SaveSetting "BOMBER2", "TOPSCORE", "NAME", "Player"
End Sub

Private Sub Form_Load()
Me.Visible = True
txtSpeed.Text = (10000 - Speed) \ 2000 'get the speed number
End Sub

Private Sub txtSpeed_Change()
'make sure the speed is in the range
If Val(txtSpeed.Text) < 1 Then txtSpeed.Text = "1"
If Val(txtSpeed.Text) > 10 Then txtSpeed.Text = "10"

Speed = 10000 - (Val(txtSpeed) * 2000) 'set the speed
SaveSetting "BOMBER2", "GENERAL", "SPEED", Speed 'saves it in the records
End Sub
