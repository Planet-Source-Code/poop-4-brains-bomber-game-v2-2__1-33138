VERSION 5.00
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bomber 2"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8970
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   292
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   598
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox boxM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   9480
      Picture         =   "frmGame.frx":08CA
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   38
      Top             =   3360
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox boxS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   9120
      Picture         =   "frmGame.frx":0DBC
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   37
      Top             =   3360
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   810
      Index           =   5
      Left            =   11160
      Picture         =   "frmGame.frx":12AE
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   36
      Top             =   2160
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   810
      Index           =   5
      Left            =   11640
      Picture         =   "frmGame.frx":18EA0
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   35
      Top             =   2280
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.PictureBox osamas 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   810
      Left            =   12240
      Picture         =   "frmGame.frx":30A92
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   34
      Top             =   2280
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox osamaM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   810
      Left            =   10080
      Picture         =   "frmGame.frx":34A1C
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   33
      Top             =   2640
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   4
      Left            =   9840
      Picture         =   "frmGame.frx":389A6
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   26
      Top             =   3840
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   4
      Left            =   9840
      Picture         =   "frmGame.frx":3D038
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   25
      Top             =   3960
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   3
      Left            =   11520
      Picture         =   "frmGame.frx":416CA
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   3
      Left            =   11520
      Picture         =   "frmGame.frx":41CFC
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   23
      Top             =   240
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   2
      Left            =   9840
      Picture         =   "frmGame.frx":4232E
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   350
      TabIndex        =   18
      Top             =   3240
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   2
      Left            =   9840
      Picture         =   "frmGame.frx":49EB8
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   350
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   9840
      Picture         =   "frmGame.frx":51A42
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   16
      Top             =   2880
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   9840
      Picture         =   "frmGame.frx":531F4
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   9840
      Picture         =   "frmGame.frx":549A6
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   9840
      Picture         =   "frmGame.frx":59038
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.PictureBox eSubm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9960
      Picture         =   "frmGame.frx":5D6CA
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox eSub 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9840
      Picture         =   "frmGame.frx":5FA34
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox eShipm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9960
      Picture         =   "frmGame.frx":61D9E
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox eShip 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9840
      Picture         =   "frmGame.frx":64108
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox bombm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   9960
      Picture         =   "frmGame.frx":66472
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox bomb 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   9840
      Picture         =   "frmGame.frx":665F4
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox eplanem 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   10320
      Picture         =   "frmGame.frx":66776
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox eplane 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   10200
      Picture         =   "frmGame.frx":68AE0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox bulletm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   9960
      Picture         =   "frmGame.frx":6AE4A
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox bullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   9840
      Picture         =   "frmGame.frx":6AEDC
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox playm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9960
      Picture         =   "frmGame.frx":6AF6E
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox play 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9840
      Picture         =   "frmGame.frx":6D2D8
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox board 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4500
      Left            =   0
      Picture         =   "frmGame.frx":6F642
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   0
      Top             =   -120
      Width           =   9000
      Begin VB.Timer tmr 
         Interval        =   1000
         Left            =   840
         Top             =   1200
      End
      Begin VB.PictureBox tscore 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   2880
         ScaleHeight     =   1335
         ScaleWidth      =   3495
         TabIndex        =   27
         Top             =   1680
         Visible         =   0   'False
         Width           =   3495
         Begin VB.CommandButton cmdOk 
            BackColor       =   &H00FF0000&
            Caption         =   "Ok"
            Height          =   255
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label lblScore 
            BackStyle       =   0  'Transparent
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1920
            TabIndex        =   32
            Top             =   240
            Width           =   1575
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
            Left            =   840
            TabIndex        =   31
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Currently Held By:"
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
            TabIndex        =   30
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label lblName 
            BackStyle       =   0  'Transparent
            Caption         =   "HotDogMan"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1920
            TabIndex        =   29
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdAbout 
         BackColor       =   &H000080FF&
         Caption         =   "About"
         Height          =   255
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdTScore 
         BackColor       =   &H000080FF&
         Caption         =   "Top Score"
         Height          =   255
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H000080FF&
         Caption         =   "New Game"
         Height          =   255
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdEnd 
         BackColor       =   &H000080FF&
         Caption         =   "Exit"
         Height          =   255
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub cmdAbout_Click()
DoCredits 'load the about screen (which is a credits screen0
End Sub

Sub cmdEnd_Click()
E.EndSound ' when ending the game you need to unload the directsound components
End 'end the whole entire app incase of open forms or the loop still going
End Sub

Sub cmdNew_Click()
Randomize 'randomize the random number thingy so it does play the same
'scene over and over again
NewGame 'start a newgame
End Sub

Function MainLoop()
Dim c As Long, AddS, AddSh, AddPl, I, AddB

MakeMenu 'set the menu captions and functions

E.StartSound 'start the sound engine (DS_ENGINE.cls)

E.LoadSound App.Path + "\explo.wav", 1 'load all of the sounds into the buffers
E.LoadSound App.Path + "\splash.wav", 2
E.LoadSound App.Path + "\shot.wav", 3
E.LoadSound App.Path + "\shot.wav", 4
E.LoadSound App.Path + "\smlexplo.wav", 5

Do
If c > Speed And GetTickCount > 500 Then 'if the delay is up and the processor delay is up
    c = 0 'reset the delay

If Pause > 0 Then 'if paused
    Pause = Pause - 1 'remove 1 from the pause time
    Running = False 'dont let the game run, but let the menu run instead
    Message = "Paused" 'set the message it doesnt look like you just started the program
    If Pause = 1 Then Running = True 'reset the game if the pause is almost up
End If

If Running = True Then 'if the game is running

If AddS > 15 Then 'if it is time to make a submarine
    AddS = 0 'reset the submarine counter
    If Int(Rnd * 5) = 3 Then MakeSub
Else
    AddS = AddS + 1
End If

If AddPl > 10 Then 'if it is time to make a plane
    AddPl = 0 'reset the plane counter
    If Int(Rnd * 5) = 3 Then MakePlane
Else
    AddPl = AddPl + 1
End If

If AddB > 10 Then 'if its time to drop a health box
    AddB = 0 'reset the box counter
    If Int(Rnd * 30) = Int(Rnd * 5) And Player.HP < (50 * 0.75) Then DropBox
Else
    AddB = AddB + 1
End If


If AddSh > 15 Then 'if its time to make a ship
    AddSh = 0 'reset the ship counter
    If Int(Rnd * 5) = 3 Then MakeShip
Else
    AddSh = AddSh + 1
End If

CheckPlayerInput 'check for the input

MoveShots 'move all objects
MoveEnemys
MovePlayer

RunExplo 'do the animation for the explosions

board.Cls

For I = 1 To 20 'draw all of the shots fired
    If PShot(I).Act = True Then
        E.DrawObj board.hdc, PShot(I).X, PShot(I).Y, 5, 5, bulletm.hdc, 0, 0, Mask
        E.DrawObj board.hdc, PShot(I).X, PShot(I).Y, 5, 5, bullet.hdc, 0, 0, sprite
    End If

    If plShot(I).Act = True Then
        E.DrawObj board.hdc, plShot(I).X, plShot(I).Y, 5, 5, bulletm.hdc, 0, 0, Mask
        E.DrawObj board.hdc, plShot(I).X, plShot(I).Y, 5, 5, bullet.hdc, 0, 0, sprite
    End If

    If ShShot(I).Act = True Then
        E.DrawObj board.hdc, ShShot(I).X, ShShot(I).Y, 5, 5, bulletm.hdc, 0, 0, Mask
        E.DrawObj board.hdc, ShShot(I).X, ShShot(I).Y, 5, 5, bullet.hdc, 0, 0, sprite
    End If

    If I < 11 Then
        If B(I).Act = True Then
            E.DrawObj board.hdc, B(I).X, B(I).Y, 10, 10, bombm.hdc, 0, 0, Mask
            E.DrawObj board.hdc, B(I).X, B(I).Y, 10, 10, bomb.hdc, 0, 0, sprite
        End If
    End If
    
Next I

For I = 1 To 10 'draw all of the enemys
    If S(I).Act = True Then
        E.DrawObj board.hdc, S(I).X, S(I).Y, 50, 30, eSubm.hdc, S(I).Dir * 50, 0, Mask
        E.DrawObj board.hdc, S(I).X, S(I).Y, 50, 30, eSub.hdc, S(I).Dir * 50, 0, sprite
    End If

    If P(I).Act = True Then
        E.DrawObj board.hdc, P(I).X, P(I).Y, 50, 30, eplanem.hdc, P(I).Dir * 50, 0, Mask
        E.DrawObj board.hdc, P(I).X, P(I).Y, 50, 30, eplane.hdc, P(I).Dir * 50, 0, sprite
    End If

    If Sh(I).Act = True Then
        E.DrawObj board.hdc, Sh(I).X, Sh(I).Y, 50, 30, eShipm.hdc, Sh(I).Dir * 50, 0, Mask
        E.DrawObj board.hdc, Sh(I).X, Sh(I).Y, 50, 30, eShip.hdc, Sh(I).Dir * 50, 0, sprite
    End If
Next I

If Osama.Act = True Then 'if osama boat is visible draw it
    E.DrawObj board.hdc, Osama.X, 203 - 45, 100, 50, osamaM.hdc, 0, 0, Mask
    E.DrawObj board.hdc, Osama.X, 203 - 45, 100, 50, osamas.hdc, 0, 0, sprite
End If

    'draw the player's aircraft
    E.DrawObj board.hdc, Player.X, Player.Y, 50, 30, playm.hdc, Player.Dir * 50, 0, Mask
    E.DrawObj board.hdc, Player.X, Player.Y, 50, 30, Play.hdc, Player.Dir * 50, 0, sprite

For I = 1 To 10 'draw all of the health boxes
    If Box(I).Act = True Then
        E.DrawObj board.hdc, Box(I).X, Box(I).Y, 20, 20, boxM.hdc, 0, 0, Mask
        E.DrawObj board.hdc, Box(I).X, Box(I).Y, 20, 20, boxS.hdc, 0, 0, sprite
    End If
Next I

For I = 1 To 20 'draw the explosions
    If Ex(I).Act = True Then
        E.DrawObj board.hdc, Ex(I).X, Ex(I).Y, Ex(I).W, Ex(I).H, explom(Ex(I).T).hdc, Ex(I).Frame * Ex(I).W, 0, Mask
        E.DrawObj board.hdc, Ex(I).X, Ex(I).Y, Ex(I).W, Ex(I).H, explo(Ex(I).T).hdc, Ex(I).Frame * Ex(I).W, 0, sprite
    End If
Next I

    'draw the score and health board up in the left hand corner
    board.CurrentX = 10
    board.CurrentY = 10
    board.Print "Score: " & Player.Score
    board.CurrentX = 10
    board.CurrentY = 20 + board.TextHeight("|")
    board.Print "HP: "
    board.Line (15 + board.TextWidth("HP: "), 20 + board.TextHeight("|"))-((15 + board.TextWidth("HP: ")) + 50, (20 + board.TextHeight("|")) + board.TextHeight("|")), vbRed, BF
    board.Line (15 + board.TextWidth("HP: "), 20 + board.TextHeight("|"))-((15 + board.TextWidth("HP: ")) + Player.HP, (20 + board.TextHeight("|")) + board.TextHeight("|")), vbGreen, BF

Else

    board.Cls 'if the game isnt running still clear the board

    'draw the big lettered message
    board.ForeColor = vbBlack
    board.FontSize = 46
    board.CurrentX = board.ScaleWidth \ 2 - board.TextWidth(Message) \ 2
    board.CurrentY = board.ScaleHeight \ 2 - (board.TextHeight("|") * 2)
    board.Print Message

    board.FontSize = 10 'reset the fontsize for drawing the menu

    'get the input for the menu
    If E.IsPressed(vbKeyDown) Then Selected = Selected + 1: If Selected > UBound(M()) Then Selected = UBound(M())
    If E.IsPressed(vbKeyUp) Then Selected = Selected - 1: If Selected < LBound(M()) Then Selected = LBound(M())
    If E.IsPressed(vbKeyControl) Then DoItem

    'draw the menu on the board
    DrawMenu board
End If

If E.IsPressed(vbKeyF3) Then 'whether or not the game is running still change the pause thing
    Select Case Pause
        Case Is <= 0: Pause = 9999999 'reset the pause so it wont run out any time soon
        Case Is > 0: Pause = 0: Running = True 'unpause the game and let the game run again
    End Select
End If

Else
    c = c + 1 'add to the delay counter
End If

    DoEvents
Loop
End Function

Private Sub cmdOK_Click() 'the ok button on the topscore box
tscore.Visible = False 'hide the topscore box
End Sub

Sub cmdTScore_Click()
tscore.Visible = True 'show the topscore box
cmdOK.SetFocus 'set focus to the ok button on the topscore box
End Sub

Private Sub Form_Load()
Speed = Val(GetSetting("BOMBER2", "GENERAL", "SPEED"))

'set the speed within the boundries
If Speed < 1000 Then Speed = 1000
If Speed > 10000 Then Speed = 10000

Message = "Bomber 2" 'set the message title
Me.Visible = True 'show the form so that the game doesnt run while the form isnt visible
MainLoop 'start the gameloop
End Sub

Private Sub Form_Resize() 'just a little something for my sloppiness
'it sets the board's coords so i dont have to do it in design mode
board.Left = 0
board.Top = 0
End Sub

Private Sub tmr_Timer()
'set the topscore box's contents (like the name and topscore)
lblScore.CAPTION = GetSetting("BOMBER2", "TOPSCORE", "SCORE")
lblName.CAPTION = GetSetting("BOMBER2", "TOPSCORE", "NAME")
End Sub
