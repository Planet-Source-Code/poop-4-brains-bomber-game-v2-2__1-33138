VERSION 5.00
Begin VB.Form frmTactics 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suggested Tactics"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6165
   Icon            =   "frmTactics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox txtTact 
      Height          =   2895
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   3495
   End
End
Attribute VB_Name = "frmTactics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Pause = 0
Unload Me
End Sub

Private Sub Form_Load()
Dim n As String, I
Me.Visible = True

txtTact.Text = ""

'load the file into the textbox
Open App.Path + "\tactics.dat" For Input As #1
lblStatus.CAPTION = "Loading Data..."

Do Until EOF(1)
    Line Input #1, n
    txtTact.Text = txtTact.Text & n & vbCrLf

    For I = 1 To 10
        DoEvents
    Next I

    DoEvents
Loop

lblStatus.CAPTION = "Data loaded"
Close #1
End Sub
