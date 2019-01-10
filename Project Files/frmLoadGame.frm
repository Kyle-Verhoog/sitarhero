VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLoadGame 
   BorderStyle     =   0  'None
   ClientHeight    =   14190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17490
   LinkTopic       =   "Form1"
   ScaleHeight     =   14190
   ScaleWidth      =   17490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer AntiCheat 
      Interval        =   500
      Left            =   13080
      Top             =   11280
   End
   Begin VB.Timer tmrRedExplosion 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   13800
      Top             =   11760
   End
   Begin VB.Timer tmrYellowExplosion 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   13800
      Top             =   12120
   End
   Begin VB.Timer tmrBlueExplosion 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   13800
      Top             =   11400
   End
   Begin VB.Timer CheckForCollisionTimer 
      Interval        =   700
      Left            =   12720
      Top             =   11880
   End
   Begin VB.PictureBox imgBackGround 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   30000
      Left            =   4080
      Picture         =   "frmLoadGame.frx":0000
      ScaleHeight     =   30000
      ScaleWidth      =   7500
      TabIndex        =   3
      Top             =   -15720
      Width           =   7500
      Begin VB.PictureBox BaseRedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Left            =   4130
         Picture         =   "frmLoadGame.frx":AF6F
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   37
         Top             =   27480
         Width           =   630
      End
      Begin VB.PictureBox BaseYellowNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Left            =   5400
         Picture         =   "frmLoadGame.frx":BF68
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   36
         Top             =   27480
         Width           =   630
      End
      Begin VB.PictureBox BaseBlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Left            =   2700
         Picture         =   "frmLoadGame.frx":CF8C
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   35
         Top             =   27480
         Width           =   630
      End
      Begin VB.PictureBox YellowNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   1
         Left            =   5400
         Picture         =   "frmLoadGame.frx":DF68
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   34
         Top             =   2520
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox RedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   1
         Left            =   4155
         Picture         =   "frmLoadGame.frx":EF8C
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   33
         Top             =   1080
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox RedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   0
         Left            =   4155
         Picture         =   "frmLoadGame.frx":FF85
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   32
         Top             =   1080
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox BlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   0
         Left            =   2760
         Picture         =   "frmLoadGame.frx":10F7E
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox YellowNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   0
         Left            =   5400
         Picture         =   "frmLoadGame.frx":11F5A
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   30
         Top             =   2520
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox BlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   1
         Left            =   2760
         Picture         =   "frmLoadGame.frx":12F7E
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.CommandButton cmdBottom 
         Caption         =   "Command1"
         Enabled         =   0   'False
         Height          =   615
         Left            =   2760
         TabIndex        =   28
         Top             =   28440
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.PictureBox YellowNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   2
         Left            =   0
         Picture         =   "frmLoadGame.frx":13F5A
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   27
         Top             =   2520
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox BlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   2
         Left            =   0
         Picture         =   "frmLoadGame.frx":14F7E
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox RedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   2
         Left            =   0
         Picture         =   "frmLoadGame.frx":15F5A
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   25
         Top             =   1080
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox YellowNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   3
         Left            =   0
         Picture         =   "frmLoadGame.frx":16F53
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   24
         Top             =   2520
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox BlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   3
         Left            =   0
         Picture         =   "frmLoadGame.frx":17F77
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox RedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   3
         Left            =   0
         Picture         =   "frmLoadGame.frx":18F53
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   22
         Top             =   1080
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox YellowNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   4
         Left            =   0
         Picture         =   "frmLoadGame.frx":19F4C
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   21
         Top             =   2520
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox BlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   4
         Left            =   0
         Picture         =   "frmLoadGame.frx":1AF70
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox RedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   4
         Left            =   0
         Picture         =   "frmLoadGame.frx":1BF4C
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   19
         Top             =   1080
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox YellowNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   5
         Left            =   2655
         Picture         =   "frmLoadGame.frx":1CF45
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   18
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox BlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   5
         Left            =   0
         Picture         =   "frmLoadGame.frx":1DF69
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   17
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox RedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   5
         Left            =   1410
         Picture         =   "frmLoadGame.frx":1EF45
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox YellowNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   6
         Left            =   2655
         Picture         =   "frmLoadGame.frx":1FF3E
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   15
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox BlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   6
         Left            =   0
         Picture         =   "frmLoadGame.frx":20F62
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   14
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox RedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   6
         Left            =   1410
         Picture         =   "frmLoadGame.frx":21F3E
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   13
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox YellowNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   7
         Left            =   2655
         Picture         =   "frmLoadGame.frx":22F37
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   12
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox BlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   7
         Left            =   0
         Picture         =   "frmLoadGame.frx":23F5B
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   11
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox RedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   7
         Left            =   1410
         Picture         =   "frmLoadGame.frx":24F37
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   10
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox YellowNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   8
         Left            =   2655
         Picture         =   "frmLoadGame.frx":25F30
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   9
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox BlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   8
         Left            =   0
         Picture         =   "frmLoadGame.frx":26F54
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   8
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox RedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   8
         Left            =   1410
         Picture         =   "frmLoadGame.frx":27F30
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   7
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox YellowNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   9
         Left            =   2655
         Picture         =   "frmLoadGame.frx":28F29
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   6
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox BlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   9
         Left            =   0
         Picture         =   "frmLoadGame.frx":29F4D
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   5
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox RedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   9
         Left            =   1410
         Picture         =   "frmLoadGame.frx":2AF29
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   4
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox imgBlueExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   900
         Index           =   1
         Left            =   2160
         Picture         =   "frmLoadGame.frx":2BF22
         ScaleHeight     =   900
         ScaleWidth      =   1890
         TabIndex        =   39
         Top             =   27240
         Visible         =   0   'False
         Width           =   1897
      End
      Begin VB.PictureBox imgRedExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   900
         Index           =   1
         Left            =   3360
         Picture         =   "frmLoadGame.frx":30CDF
         ScaleHeight     =   900
         ScaleWidth      =   1890
         TabIndex        =   38
         Top             =   27240
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.PictureBox imgYellowExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   900
         Index           =   1
         Left            =   4800
         Picture         =   "frmLoadGame.frx":35A9C
         ScaleHeight     =   900
         ScaleWidth      =   1890
         TabIndex        =   40
         Top             =   27240
         Visible         =   0   'False
         Width           =   1897
      End
      Begin VB.PictureBox imgRedExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   900
         Index           =   2
         Left            =   3480
         Picture         =   "frmLoadGame.frx":3A859
         ScaleHeight     =   900
         ScaleWidth      =   1890
         TabIndex        =   42
         Top             =   27225
         Visible         =   0   'False
         Width           =   1897
      End
      Begin VB.PictureBox imgBlueExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   900
         Index           =   2
         Left            =   2160
         Picture         =   "frmLoadGame.frx":3FB45
         ScaleHeight     =   900
         ScaleWidth      =   1890
         TabIndex        =   41
         Top             =   27225
         Visible         =   0   'False
         Width           =   1897
      End
      Begin VB.PictureBox imgYellowExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   900
         Index           =   2
         Left            =   4800
         Picture         =   "frmLoadGame.frx":44E31
         ScaleHeight     =   900
         ScaleWidth      =   1890
         TabIndex        =   43
         Top             =   27225
         Visible         =   0   'False
         Width           =   1897
      End
   End
   Begin VB.Timer MoveNotes 
      Interval        =   20
      Left            =   10320
      Tag             =   " "
      Top             =   4440
   End
   Begin VB.Timer Game 
      Interval        =   30
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer tmrDisplayAnimation 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2760
      Top             =   4200
   End
   Begin VB.Timer tmrSpecial 
      Interval        =   30
      Left            =   13320
      Top             =   11880
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblStreakDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   1320
      TabIndex        =   0
      Top             =   10200
      Width           =   1695
   End
   Begin VB.Shape shpYellow 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008080&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   11760
      Shape           =   3  'Circle
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Shape shpRed 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   11760
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Shape shpGreen 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   11760
      Shape           =   3  'Circle
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Shape shpBack 
      FillStyle       =   0  'Solid
      Height          =   4095
      Left            =   11640
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label lblScoreDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   11520
      TabIndex        =   2
      Top             =   10080
      Width           =   1815
   End
   Begin VB.Label lblDisplay 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   3735
   End
   Begin VB.Image lblStreakBack 
      Height          =   4605
      Left            =   0
      Picture         =   "frmLoadGame.frx":4A11D
      Top             =   7920
      Width           =   4110
   End
   Begin VB.Image imgNeutral 
      Enabled         =   0   'False
      Height          =   18000
      Left            =   0
      Picture         =   "frmLoadGame.frx":4EF46
      Top             =   0
      Width           =   22500
   End
   Begin VB.Image imgHalfFail 
      Enabled         =   0   'False
      Height          =   18000
      Left            =   -120
      Picture         =   "frmLoadGame.frx":7DF56
      Top             =   -120
      Visible         =   0   'False
      Width           =   22500
   End
   Begin VB.Image imgHalfWin 
      Enabled         =   0   'False
      Height          =   18000
      Left            =   0
      Picture         =   "frmLoadGame.frx":BE617
      Top             =   0
      Visible         =   0   'False
      Width           =   22500
   End
   Begin VB.Image imgWinning 
      Enabled         =   0   'False
      Height          =   18000
      Left            =   0
      Picture         =   "frmLoadGame.frx":E4C43
      Top             =   0
      Visible         =   0   'False
      Width           =   22500
   End
   Begin VB.Image imgFail 
      Enabled         =   0   'False
      Height          =   18000
      Left            =   0
      Picture         =   "frmLoadGame.frx":123187
      Top             =   0
      Visible         =   0   'False
      Width           =   22500
   End
End
Attribute VB_Name = "frmLoadGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sitar Hero 2011
'Written by Ahmad J, Aruran R, Kyle V and Saad M
'Due January 23 2012
Dim Space As Boolean
Dim F1, F2, F3, F4 As Boolean
Dim Difficulty As Integer
Dim UpperBound, LowerBound As Integer
Dim RandomPosition As Integer
Dim FileContents As String
Dim SpaceCounter As Integer
Dim HitNote As Boolean
Dim MissNote As Boolean
'Used for Blue note animations
Dim Counter1 As Integer
'Used for Red note animations
Dim Counter2 As Integer
'Used for Yellow note animations
Dim Counter3 As Integer
Dim FolderName As String
Dim SongName As String
Dim NoteSpeed As String
Dim position As Integer

Private Sub CheckForCollision()
For Index = 0 To NumberofNotes
    If BlueNote(Index).Visible = True Then
        If Collision(BlueNote(Index), cmdBottom) = True Then
            BlueNote(Index).Visible = False
            'Random placing just to make sure the note does not interfere with anything else
            BlueNote(Index).Left = 7080
            BlueNote(Index).Top = 7560
        End If
    End If
    If RedNote(Index).Visible = True Then
        If Collision(RedNote(Index), cmdBottom) = True Then
            RedNote(Index).Visible = False
            'Random placing just to make sure the note does not interfere with anything else
            RedNote(Index).Left = 7080
            RedNote(Index).Top = 8640
        End If
    End If
    If YellowNote(Index).Visible = True Then
        If Collision(YellowNote(Index), cmdBottom) = True Then
            YellowNote(Index).Visible = False
            'Random placing just to make sure the note does not interfere with anything else
            YellowNote(Index).Left = 7080
            YellowNote(Index).Top = 9960
        End If
    End If
Next



For Index = 0 To NumberofNotes
    If BlueNote(Index).Visible = True Then
        If Collision(BlueNote(Index), cmdBottom) = True Then
            WriteMissedNote
            Score = Score - Difficulty
            MissNote = True
            HitNote = False
            BlueNote(Index).Visible = False
            If Notes > 0 Then
                Notes = 0
            Else
                Notes = Notes - 1
            End If
            End If
    End If
    If RedNote(Index).Visible = True Then
        If Collision(RedNote(Index), cmdBottom) = True Then
            WriteMissedNote
            Score = Score - Difficulty
            MissNote = True
            HitNote = False
            RedNote(Index).Visible = False
            If Notes > 0 Then
                Notes = 0
            Else
                Notes = Notes - 1
            End If
        End If
    End If
    If YellowNote(Index).Visible = True Then
        If Collision(YellowNote(Index), cmdBottom) = True Then
            WriteMissedNote
            Score = Score - Difficulty
            MissNote = True
            HitNote = False
            YellowNote(Index).Visible = False
            If Notes > 0 Then
                Notes = 0
            Else
                Notes = Notes - 1
            End If
        End If
    End If
Next
Select Case Notes
    Case Is < -30
        Difficulty = 200
    Case Is < -20
        Difficulty = 100
        'frmGame.BackColor = &HFF&
    Case Is < -15
        Difficulty = 70
    Case Is < -10
        Difficulty = 60
    Case Is < -5
        Difficulty = 50
    Case Is < -3
        Difficulty = 40
    Case Is > 50
        Cheer
        'frmGame.BackColor = &HFF0000
        Multiplier = 20
    Case Is > 20
        Multiplier = 15
        'frmGame.BackColor = &HFF00&
    Case Else
        frmGame.BackColor = &H80FF&
        Difficulty = 30
        Multiplier = 10
End Select
If HitNote = True Then
    MissNote = False
End If
If MissNote = True Then
    
End If
End Sub

Private Sub AntiCheat_Timer()
SpaceCounter = 0
AntiCheat.Enabled = False
End Sub
Private Sub cmdQuit_Click()
cmdQuit.Enabled = True
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'37 Left Arrow
'38 Up Arrow
'39 Right Arrow
'40 Down Arrow
'32 Space
'112 F1
'113 F2
'114 F3
'115 F4
Select Case KeyCode
    Case 112
        F1 = True
        BaseBlueNote.Picture = LoadPicture(App.Path & "\BlueNoteBase.JPG")
    Case 113
        F2 = True
        BaseRedNote.Picture = LoadPicture(App.Path & "\RedNoteBase.JPG")
    Case 114
        F3 = True
        BaseYellowNote.Picture = LoadPicture(App.Path & "\YellowNoteBase.JPG")
    Case 115
        F4 = True
    Case 49
        F1 = True
        BaseBlueNote.Picture = LoadPicture(App.Path & "\BlueNoteBase.JPG")
    Case 50
        F2 = True
        BaseRedNote.Picture = LoadPicture(App.Path & "\RedNoteBase.JPG")
    Case 51
        F3 = True
        BaseYellowNote.Picture = LoadPicture(App.Path & "\YellowNoteBase.JPG")
    Case 40
        Space = True
    Case 32
        Space = True
        SpaceCounter = SpaceCounter + 1
        AntiCheat.Enabled = True
    Case 76
        Unload Me
    Case 27
        Unload Me
    Case 80
        MsgBox "Press Ok to Resume"
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 32
End Select

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 112
        F1 = False
        BaseBlueNote.Picture = LoadPicture(App.Path & "\BlueNote.JPG")
    Case 113
        F2 = False
        BaseRedNote.Picture = LoadPicture(App.Path & "\RedNote.JPG")
    Case 114
        F3 = False
        BaseYellowNote.Picture = LoadPicture(App.Path & "\YellowNote.JPG")
    Case 115
        F4 = False
    Case 49
        F1 = False
        BaseBlueNote.Picture = LoadPicture(App.Path & "\BlueNote.JPG")
    Case 50
        F2 = False
        BaseRedNote.Picture = LoadPicture(App.Path & "\RedNote.JPG")
    Case 51
        F3 = False
        BaseYellowNote.Picture = LoadPicture(App.Path & "\YellowNote.JPG")
    Case 40
        Space = False
    Case 32
        Space = False
End Select
End Sub

Private Sub Form_Load()
On Error Resume Next
FolderName = InputBox("Enter Song Folder Name")

Open App.Path & "\" & FolderName & "\Song.txt" For Input As #1
Input #1, SongName
Close #1

Open App.Path & "\" & FolderName & "\Difficulty.txt" For Input As #1
Input #1, NoteSpeed
Close #1

Select Case NoteSpeed
Case "Beginner"
    MoveNotes.Interval = 27
Case "Easy"
    MoveNotes.Interval = 25
Case "Medium"
    MoveNotes.Interval = 20
Case "Hard"
    MoveNotes.Interval = 12
Case "Impossible"
    MoveNotes.Interval = 7
End Select
Open App.Path & "\" & FolderName & "\Song.txt" For Output As #1
Print #1, SongName;
Close #1

Open App.Path & "\" & FolderName & "\NumberOfNotes.txt" For Input As #1
Input #1, NumberofNotes
Close #1

Open App.Path & "\Cheer-Boo.txt" For Output As #1
Print #1, ""
Close #1

Open App.Path & "\MissedNote.txt" For Output As #1
Print #1, "hit"
Close #1

Open App.Path & "\Command.txt" For Output As #1
Print #1, "start";
Close #1

Open App.Path & "\Command.txt" For Output As #1
Print #1, "playing";
Close #1

    For BlueIndex = 0 To NumberofNotes - 1 Step 1
        BlueNote(BlueIndex).Visible = True
        Open App.Path & "\" & FolderName & "\BlueNote" & BlueIndex & ".note" For Input As #1
        Input #1, position
        Close #1
        BlueNote(BlueIndex).Left = 2700
        BlueNote(BlueIndex).Top = position
    Next
    For RedIndex = 0 To NumberofNotes - 1 Step 1
        RedNote(RedIndex).Visible = True
        Open App.Path & "\" & FolderName & "\RedNote" & RedIndex & ".note" For Input As #2
        Input #2, position
        Close #2
        RedNote(RedIndex).Left = 4130
        RedNote(RedIndex).Top = position
    Next
    For YellowIndex = 0 To NumberofNotes - 1 Step 1
        YellowNote(YellowIndex).Visible = True
        Open App.Path & "\" & FolderName & "\YellowNote" & YellowIndex & ".note" For Input As #3
        Input #3, position
        Close #3
        YellowNote(YellowIndex).Left = 5400
        YellowNote(YellowIndex).Top = position
    Next
HitNote = True
Score = 1000
Difficulty = 30
Multiplier = 10
End Sub

Private Sub Form_Terminate()
    Open App.Path & "\MissedNote.txt" For Output As #1
            Print #1, "exit"
            Close #1
End Sub

Private Sub Game_Timer()
CheckForCollision
For Index = 0 To NumberofNotes
If SpaceCounter = 1 Then
    If Space = True And F1 = True And F2 = True And F3 = True Then
        If Collision(BlueNote(Index), BaseBlueNote) And Collision(RedNote(Index), BaseRedNote) And Collision(YellowNote(Index), BaseYellowNote) Then
            Call HitNoteFunction(Multiplier, 1, BlueNote(Index))
            Call HitNoteFunction(Multiplier, 1, RedNote(Index))
            Call HitNoteFunction(Multiplier, 1, YellowNote(Index))
            Counter1 = 0
            tmrBlueExplosion.Enabled = True
            Counter2 = 0
            tmrRedExplosion.Enabled = True
            Counter3 = 0
            tmrYellowExplosion.Enabled = True
            HitNote = True
            SpaceCounter = 0
        End If
    End If
    If Space = True And F1 = True And F2 = True And F3 = False Then
        If Collision(BlueNote(Index), BaseBlueNote) And Collision(RedNote(Index), BaseRedNote) Then
            Call HitNoteFunction(Multiplier, 1, BlueNote(Index))
            Call HitNoteFunction(Multiplier, 1, RedNote(Index))
            Counter1 = 0
            tmrBlueExplosion.Enabled = True
            Counter2 = 0
            tmrRedExplosion.Enabled = True
            HitNote = True
            SpaceCounter = 0
        End If
    End If
    If Space = True And F1 = False And F2 = True And F3 = True Then
        If Collision(YellowNote(Index), BaseYellowNote) And Collision(RedNote(Index), BaseRedNote) Then
            Call HitNoteFunction(Multiplier, 1, RedNote(Index))
            Call HitNoteFunction(Multiplier, 1, YellowNote(Index))
            Counter2 = 0
            tmrRedExplosion.Enabled = True
            Counter3 = 0
            tmrYellowExplosion.Enabled = True
            HitNote = True
            SpaceCounter = 0
        End If
    End If
    If Space = True And F1 = True And F2 = False And F3 = True Then
        If Collision(BlueNote(Index), BaseBlueNote) And Collision(YellowNote(Index), BaseYellowNote) Then
            Call HitNoteFunction(Multiplier, 1, BlueNote(Index))
            Call HitNoteFunction(Multiplier, 1, YellowNote(Index))
            Counter1 = 0
            tmrBlueExplosion.Enabled = True
            Counter3 = 0
            tmrYellowExplosion.Enabled = True
            HitNote = True
            SpaceCounter = 0
        End If
    End If
    If Space = True And F1 = True And F2 = False And F3 = False Then
        If Collision(BlueNote(Index), BaseBlueNote) Then
            Call HitNoteFunction(Multiplier, 1, BlueNote(Index))
            Counter1 = 0
            tmrBlueExplosion.Enabled = True
            HitNote = True
            SpaceCounter = 0
        End If
    End If
    If Space = True And F1 = False And F2 = True And F3 = False Then
        If Collision(RedNote(Index), BaseRedNote) Then
            Call HitNoteFunction(Multiplier, 1, RedNote(Index))
            Counter2 = 0
            tmrRedExplosion.Enabled = True
            HitNote = True
            SpaceCounter = 0
        End If
    End If
    If Space = True And F1 = False And F2 = False And F3 = True Then
        If Collision(YellowNote(Index), BaseYellowNote) Then
            Call HitNoteFunction(Multiplier, 1, YellowNote(Index))
            Counter3 = 0
            tmrYellowExplosion.Enabled = True
            HitNote = True
            SpaceCounter = 0
        End If
    End If
End If
Next



If Score <= 0 Then
Open App.Path & "\Cheer-Boo.txt" For Output As #1
            Print #1, "boo"
            Close #1
Game.Enabled = False
    MsgBox "You got booed off stage!"
End If

Select Case Score
    Case Is <= 500
        Boo
        imgFail.Visible = True
        imgHalfFail.Visible = False
        shpRed.FillColor = &HFF&
        shpGreen.FillColor = &H8000&
        shpYellow.FillColor = &H8080&
    Case Is <= 700
        imgHalfFail.Visible = True
        imgNeutral.Visible = False
        imgHalfWin.Visible = False
        imgFail.Visible = False
        shpYellow.FillColor = &HFFFF&
        shpRed.FillColor = &H80&
    Case Is >= 2000
        Cheer
        imgWinning.Visible = True
        imgHalfWin.Visible = False
    Case Is >= 1200
        Cheer
        imgHalfWin.Visible = True
        imgNeutral.Visible = False
        imgWinning.Visible = False
        shpGreen.FillColor = &HFF00&
        shpRed.FillColor = &H80&
        shpYellow.FillColor = &H8080&
    Case Else
        imgNeutral.Visible = True
        shpYellow.FillColor = &HFFFF&
        shpRed.FillColor = &H80&
        shpGreen.FillColor = &H8000&
End Select
'Actions Related to the note streak------------------------------------------
If Notes >= 30 And Notes <= 36 Then
    tmrDisplayAnimation.Enabled = True
    lblDisplay.Caption = "30 Note Streak"
ElseIf Notes >= 20 And Notes <= 26 Then
    tmrDisplayAnimation.Enabled = True
    lblDisplay.Caption = "20 Note Streak"
End If

Select Case Notes
        
    Case Is >= 20
        lblStreakDisplay.ForeColor = &HFF00&
    Case Is >= 10
        lblStreakDisplay.ForeColor = &HFFFF&
    Case Else
        lblStreakDisplay.ForeColor = &HFF&
End Select
CheckForCollisionTimer.Enabled = True
Space = False
If Notes < 0 Then
    lblStreakDisplay.Caption = 0
Else
    lblStreakDisplay.Caption = Notes
End If

lblScoreDisplay.Caption = Score
End Sub

Private Sub MoveNotes_Timer()
For Index = 0 To NumberofNotes
BlueNote(Index).Top = BlueNote(Index).Top + 50
RedNote(Index).Top = RedNote(Index).Top + 50
YellowNote(Index).Top = YellowNote(Index).Top + 50
Next
End Sub


Private Sub tmrBlueExplosion_Timer()
    BaseBlueNote.Visible = False
    Select Case Counter1
    Case Is = 0
        imgBlueExplode(1).Visible = False
        imgBlueExplode(2).Visible = False
    Case Is = 1
        imgBlueExplode(1).Visible = True
        imgBlueExplode(2).Visible = False
    Case Is = 2
        imgBlueExplode(1).Visible = False
        imgBlueExplode(2).Visible = True
    Case Else
        BaseBlueNote.Visible = True
        imgBlueExplode(1).Visible = False
        imgBlueExplode(2).Visible = False
        tmrBlueExplosion.Enabled = False
    End Select
    Counter1 = Counter1 + 1
End Sub

Private Sub tmrDisplayAnimation_Timer()
lblDisplay.Visible = True
lblDisplay.FontSize = lblDisplay.FontSize + 0.5
If lblDisplay.FontSize > 30 Then
    tmrDisplayAnimation.Enabled = False
    lblDisplay.Visible = False
    lblDisplay.FontSize = 24
End If
End Sub


Private Sub tmrRedExplosion_Timer()
    BaseRedNote.Visible = False
    Select Case Counter2
    Case Is = 0
        imgRedExplode(1).Visible = False
        imgRedExplode(2).Visible = False
    Case Is = 1
        imgRedExplode(1).Visible = True
        imgRedExplode(2).Visible = False
    Case Is = 2
        imgRedExplode(1).Visible = False
        imgRedExplode(2).Visible = True
    Case Else
        BaseRedNote.Visible = True
        Counter = 0
        imgRedExplode(1).Visible = False
        imgRedExplode(2).Visible = False
        tmrRedExplosion.Enabled = False
    End Select
    Counter2 = Counter2 + 1
End Sub

Private Sub tmrYellowExplosion_Timer()
    BaseYellowNote.Visible = False
    Select Case Counter3
    Case Is = 0
        imgYellowExplode(1).Visible = False
        imgYellowExplode(2).Visible = False
    Case Is = 1
        imgYellowExplode(1).Visible = True
        imgYellowExplode(2).Visible = False
    Case Is = 2
        imgYellowExplode(1).Visible = False
        imgYellowExplode(2).Visible = True
    Case Else
        BaseYellowNote.Visible = True
        imgYellowExplode(1).Visible = False
        imgYellowExplode(2).Visible = False
        tmrYellowExplosion.Enabled = False
    End Select
    Counter3 = Counter3 + 1
End Sub

