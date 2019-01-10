VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSongEditor 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Song Editor"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog SelectSong 
      Left            =   8040
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fmeSpeed 
      Caption         =   "Speed"
      Height          =   2175
      Left            =   5640
      TabIndex        =   45
      Top             =   6240
      Width           =   1455
      Begin VB.OptionButton optImpossible 
         Caption         =   "Impossible"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton optHard 
         Caption         =   "Hard"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optMedium 
         Caption         =   "Medium"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton optEasy 
         Caption         =   "Easy"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optBeginner 
         Caption         =   "Beginner"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Timer tmrResetScrollbar 
      Interval        =   1
      Left            =   5880
      Top             =   3720
   End
   Begin VB.Timer tmrBlueExplosion 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   5640
      Top             =   8640
   End
   Begin VB.Timer AntiCheat 
      Interval        =   500
      Left            =   7680
      Top             =   9120
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5055
      Left            =   5520
      Max             =   30000
      TabIndex        =   3
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton cmdAddYellow 
      Caption         =   "Add Yellow Note"
      Height          =   615
      Left            =   8040
      TabIndex        =   2
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddRed 
      Caption         =   "Add Red Note"
      Height          =   615
      Left            =   8040
      TabIndex        =   1
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddBlue 
      Caption         =   "Add Blue Note"
      Height          =   615
      Left            =   8040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog Song 
      Left            =   8520
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrYellowExplosion 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   5640
      Top             =   9360
   End
   Begin VB.Timer tmrRedExplosion 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   5640
      Top             =   9000
   End
   Begin VB.Timer tmrSpecial 
      Interval        =   30
      Left            =   13200
      Top             =   11760
   End
   Begin VB.Timer tmrDisplayAnimation 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8040
      Top             =   9120
   End
   Begin VB.Timer CheckForCollisionTimer 
      Interval        =   700
      Left            =   8400
      Top             =   9120
   End
   Begin VB.Timer Game 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   7080
      Top             =   120
   End
   Begin VB.Timer MoveNotes 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   6000
      Tag             =   " "
      Top             =   2400
   End
   Begin VB.PictureBox imgBackGround 
      Enabled         =   0   'False
      Height          =   30000
      Left            =   -1920
      Picture         =   "frmSongEditor.frx":0000
      ScaleHeight     =   29940
      ScaleWidth      =   7440
      TabIndex        =   4
      Top             =   -20160
      Width           =   7500
      Begin VB.PictureBox RedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   9
         Left            =   1410
         Picture         =   "frmSongEditor.frx":AF6F
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   44
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox BlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   9
         Left            =   0
         Picture         =   "frmSongEditor.frx":BF68
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   43
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox YellowNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   9
         Left            =   2655
         Picture         =   "frmSongEditor.frx":CF44
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   42
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox RedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   8
         Left            =   1410
         Picture         =   "frmSongEditor.frx":DF68
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   41
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox BlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   8
         Left            =   0
         Picture         =   "frmSongEditor.frx":EF61
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   40
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox YellowNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   8
         Left            =   2655
         Picture         =   "frmSongEditor.frx":FF3D
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   39
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox RedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   7
         Left            =   1410
         Picture         =   "frmSongEditor.frx":10F61
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   38
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox BlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   7
         Left            =   0
         Picture         =   "frmSongEditor.frx":11F5A
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   37
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox YellowNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   7
         Left            =   2655
         Picture         =   "frmSongEditor.frx":12F36
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   36
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox RedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   6
         Left            =   1410
         Picture         =   "frmSongEditor.frx":13F5A
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   35
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox BlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   6
         Left            =   0
         Picture         =   "frmSongEditor.frx":14F53
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   34
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox YellowNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   6
         Left            =   2655
         Picture         =   "frmSongEditor.frx":15F2F
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   33
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox RedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   5
         Left            =   1410
         Picture         =   "frmSongEditor.frx":16F53
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   32
         Top             =   0
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox BlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   5
         Left            =   0
         Picture         =   "frmSongEditor.frx":17F4C
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   31
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox YellowNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   5
         Left            =   2655
         Picture         =   "frmSongEditor.frx":18F28
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   30
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox RedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   4
         Left            =   0
         Picture         =   "frmSongEditor.frx":19F4C
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   29
         Top             =   1080
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox BlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   4
         Left            =   0
         Picture         =   "frmSongEditor.frx":1AF45
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox YellowNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   4
         Left            =   0
         Picture         =   "frmSongEditor.frx":1BF21
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   27
         Top             =   2520
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox RedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   3
         Left            =   0
         Picture         =   "frmSongEditor.frx":1CF45
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   26
         Top             =   1080
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox BlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   3
         Left            =   0
         Picture         =   "frmSongEditor.frx":1DF3E
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox YellowNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   3
         Left            =   0
         Picture         =   "frmSongEditor.frx":1EF1A
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   24
         Top             =   2520
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox RedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   2
         Left            =   0
         Picture         =   "frmSongEditor.frx":1FF3E
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   23
         Top             =   1080
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox BlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   2
         Left            =   0
         Picture         =   "frmSongEditor.frx":20F37
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox YellowNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   2
         Left            =   0
         Picture         =   "frmSongEditor.frx":21F13
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   21
         Top             =   2520
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.CommandButton cmdBottom 
         Caption         =   "Command1"
         Enabled         =   0   'False
         Height          =   615
         Left            =   2760
         TabIndex        =   17
         Top             =   29520
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.PictureBox BlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   1
         Left            =   2760
         Picture         =   "frmSongEditor.frx":22F37
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
         Index           =   0
         Left            =   5400
         Picture         =   "frmSongEditor.frx":23F13
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   15
         Top             =   2520
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox BlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   0
         Left            =   2760
         Picture         =   "frmSongEditor.frx":24F37
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox RedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   0
         Left            =   4155
         Picture         =   "frmSongEditor.frx":25F13
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox RedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   1
         Left            =   4155
         Picture         =   "frmSongEditor.frx":26F0C
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox YellowNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   1
         Left            =   5400
         Picture         =   "frmSongEditor.frx":27F05
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   11
         Top             =   2520
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox BaseBlueNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Left            =   2700
         Picture         =   "frmSongEditor.frx":28F29
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   7
         Top             =   28440
         Width           =   630
      End
      Begin VB.PictureBox BaseYellowNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Left            =   5400
         Picture         =   "frmSongEditor.frx":29F05
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   6
         Top             =   28440
         Width           =   630
      End
      Begin VB.PictureBox BaseRedNote 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Left            =   4130
         Picture         =   "frmSongEditor.frx":2AF29
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   5
         Top             =   28440
         Width           =   630
      End
      Begin VB.PictureBox imgYellowExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   900
         Index           =   1
         Left            =   4680
         Picture         =   "frmSongEditor.frx":2BF22
         ScaleHeight     =   900
         ScaleWidth      =   1890
         TabIndex        =   10
         Top             =   28080
         Visible         =   0   'False
         Width           =   1897
      End
      Begin VB.PictureBox imgRedExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   900
         Index           =   1
         Left            =   3480
         Picture         =   "frmSongEditor.frx":30CDF
         ScaleHeight     =   900
         ScaleWidth      =   1890
         TabIndex        =   8
         Top             =   28080
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.PictureBox imgBlueExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   900
         Index           =   1
         Left            =   1920
         Picture         =   "frmSongEditor.frx":35A9C
         ScaleHeight     =   900
         ScaleWidth      =   1890
         TabIndex        =   9
         Top             =   28080
         Visible         =   0   'False
         Width           =   1897
      End
      Begin VB.PictureBox imgRedExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   900
         Index           =   0
         Left            =   3360
         Picture         =   "frmSongEditor.frx":3A859
         ScaleHeight     =   900
         ScaleWidth      =   1890
         TabIndex        =   51
         Top             =   28170
         Visible         =   0   'False
         Width           =   1897
      End
      Begin VB.PictureBox imgBlueExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   900
         Index           =   0
         Left            =   2040
         Picture         =   "frmSongEditor.frx":3FB45
         ScaleHeight     =   900
         ScaleWidth      =   1890
         TabIndex        =   52
         Top             =   28170
         Visible         =   0   'False
         Width           =   1897
      End
      Begin VB.PictureBox imgYellowExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   900
         Index           =   0
         Left            =   4680
         Picture         =   "frmSongEditor.frx":44E31
         ScaleHeight     =   900
         ScaleWidth      =   1890
         TabIndex        =   53
         Top             =   28080
         Visible         =   0   'False
         Width           =   1897
      End
      Begin VB.PictureBox imgRedExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   900
         Index           =   3
         Left            =   3360
         Picture         =   "frmSongEditor.frx":4A11D
         ScaleHeight     =   900
         ScaleWidth      =   1890
         TabIndex        =   54
         Top             =   28170
         Visible         =   0   'False
         Width           =   1897
      End
      Begin VB.PictureBox imgBlueExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   900
         Index           =   3
         Left            =   2040
         Picture         =   "frmSongEditor.frx":4F409
         ScaleHeight     =   900
         ScaleWidth      =   1890
         TabIndex        =   55
         Top             =   28170
         Visible         =   0   'False
         Width           =   1897
      End
      Begin VB.PictureBox imgYellowExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   900
         Index           =   3
         Left            =   4680
         Picture         =   "frmSongEditor.frx":546F5
         ScaleHeight     =   900
         ScaleWidth      =   1890
         TabIndex        =   56
         Top             =   28080
         Visible         =   0   'False
         Width           =   1897
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox imgYellowExplode 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   900
      Index           =   2
      Left            =   2880
      Picture         =   "frmSongEditor.frx":599E1
      ScaleHeight     =   900
      ScaleWidth      =   1890
      TabIndex        =   20
      Top             =   8040
      Visible         =   0   'False
      Width           =   1897
   End
   Begin VB.PictureBox imgRedExplode 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   900
      Index           =   2
      Left            =   1560
      Picture         =   "frmSongEditor.frx":5ECCD
      ScaleHeight     =   900
      ScaleWidth      =   1890
      TabIndex        =   19
      Top             =   8040
      Visible         =   0   'False
      Width           =   1897
   End
   Begin VB.PictureBox imgBlueExplode 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   900
      Index           =   2
      Left            =   120
      Picture         =   "frmSongEditor.frx":63FB9
      ScaleHeight     =   900
      ScaleWidth      =   1890
      TabIndex        =   18
      Top             =   8040
      Visible         =   0   'False
      Width           =   1897
   End
   Begin VB.Shape shpRed 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   120
      Width           =   1335
   End
   Begin VB.Shape shpGreen 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Shape shpBack 
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   8040
      Top             =   0
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpenSong 
         Caption         =   "Open Project"
      End
      Begin VB.Menu mnuSelectSong 
         Caption         =   "Select Song"
      End
      Begin VB.Menu mnuSaveSong 
         Caption         =   "Save Project"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuInsertNote 
         Caption         =   "Insert Note"
         Begin VB.Menu mnuTrack1 
            Caption         =   "Track #1"
         End
         Begin VB.Menu mnuTrack2 
            Caption         =   "Track #2"
         End
         Begin VB.Menu mnuTrack3 
            Caption         =   "Track #3"
         End
      End
   End
   Begin VB.Menu mnuTest 
      Caption         =   "Test"
      Begin VB.Menu mnuPlay 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop"
      End
   End
End
Attribute VB_Name = "frmSongEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sitar Hero 2011
'Written by Ahmad J, Aruran R, Kyle V and Saad M
'Due January 23 2012
'Sitar Hero is a guitar-hero like game in which the objective is to hit the oncoming notes
Dim Space As Boolean
Dim F1, F2, F3, F4 As Boolean
Dim Difficulty As Integer
Dim UpperBound, LowerBound As Integer
Dim RandomPosition As Integer
Dim FileContents As String
Dim SpaceCounter As Integer
Dim HitNote As Boolean
Dim MissNote As Boolean
Dim AddBlue As Boolean
Dim AddRed As Boolean
Dim AddYellow As Boolean
'Variables used for different note arrays
Dim BlueIndex As Integer
Dim RedIndex As Integer
Dim YellowIndex As Integer
'Used for Blue note animations
Dim Counter1 As Integer
'Used for Red note animations
Dim Counter2 As Integer
'Used for Yellow note animations
Dim Counter3 As Integer
Const BlueLine As Integer = 2700
Const RedLine As Integer = 4130
Const YellowLine As Integer = 5400
Const Beginner As Integer = 27
Const Easy As Integer = 25
Const Medium As Integer = 20
Const Hard As Integer = 12
Const Impossible As Integer = 7
Dim Distance As Integer
'stores the difficulty
Dim NoteSpeed As String
Dim Saved As Boolean
'Used for returns from functions
Dim ReturnVal As String
Dim SongName As String



Private Sub CheckForCollision()
For Index = 0 To NumberofNotes + 1
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
End Sub

Private Sub cmdAddBlue_Click()
AddBlue = True
BlueNote(BlueIndex).Top = imgBackGround.Top
BlueNote(BlueIndex).Visible = True
End Sub

Private Sub cmdAddRed_Click()
AddRed = True
RedNote(RedIndex).Top = imgBackGround.Top
RedNote(RedIndex).Visible = True
End Sub

Private Sub cmdAddYellow_Click()
AddYellow = True
YellowNote(YellowIndex).Top = imgBackGround.Top
YellowNote(YellowIndex).Visible = True
End Sub

Private Sub Form_Click()
If AddBlue = True Then
    If Saved = True Then
        Open Song.FileName & "\BlueNote" & BlueIndex & ".note" For Output As #1
        Print #1, BlueNote(BlueIndex).Top
        Close #1
    Else
        Save
        Open Song.FileName & "\BlueNote" & BlueIndex & ".note" For Output As #1
        Print #1, BlueNote(BlueIndex).Top
        Close #1
        Saved = True
    End If
    BlueIndex = BlueIndex + 1
    AddBlue = False
End If

If AddRed = True Then
    If Saved = True Then
        Open Song.FileName & "\RedNote" & RedIndex & ".note" For Output As #1
        Print #1, RedNote(RedIndex).Top
        Close #1
    Else
        Save
        Open Song.FileName & "\RedNote" & RedIndex & ".note" For Output As #1
        Print #1, RedNote(RedIndex).Top
        Close #1
        Saved = True
    End If
    RedIndex = RedIndex + 1
    AddRed = False
End If

If AddYellow = True Then
    If Saved = True Then
        Open Song.FileName & "\YellowNote" & YellowIndex & ".note" For Output As #1
        Print #1, YellowNote(YellowIndex).Top
        Close #1
    Else
        Save
        Open Song.FileName & "\YellowNote" & YellowIndex & ".note" For Output As #1
        Print #1, YellowNote(YellowIndex).Top
        Close #1
        Saved = True
    End If
    YellowIndex = YellowIndex + 1
    NumberofNotes = NumberofNotes + 1
    AddYellow = False
End If
End Sub

Private Sub Form_Load()
Saved = False
'Always one less because of array (0) etc 10 notes would be 9
NumberofNotes = 0
BlueIndex = NumberofNotes
RedIndex = NumberofNotes
YellowIndex = NumberofNotes
For Index = 0 To 9
BlueNote(Index).Visible = False
RedNote(Index).Visible = False
YellowNote(Index).Visible = False

Next
HitNote = True
UpperBound = -600
LowerBound = -10
Score = 1000
Difficulty = 30
Multiplier = 10
VScroll1.Value = 20200
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If AddBlue = True Then
    BlueNote(BlueIndex).Left = BlueLine
        BlueNote(BlueIndex).Top = Y + Distance
 End If
 If AddRed = True Then
    RedNote(RedIndex).Left = RedLine
        RedNote(RedIndex).Top = Y + Distance
 End If
 If AddYellow = True Then
    YellowNote(YellowIndex).Left = YellowLine
    YellowNote(YellowIndex).Top = Y + Distance
 End If
End Sub

Private Sub mnuPlay_Click()
Open App.Path & "\MissedNote.txt" For Output As #1
Print #1, "hit"
Close #1
Open App.Path & "\Command.txt" For Output As #1
Print #1, "start"
Close #1
    For Index = 1 To 20200 Step 20
        VScroll1.Value = Index
    Next
Open Song.FileName & "\Difficulty.txt" For Output As #1
Print #1, NoteSpeed
Close #1

Open Song.FileName & "\NumberOfNotes.txt" For Output As #1
Print #1, NumberofNotes
Close #1

shpGreen.FillColor = &HFF00&
shpRed.FillColor = &H80&
fmeSpeed.Enabled = False
cmdAddRed.Enabled = False
cmdAddBlue.Enabled = False
cmdAddYellow.Enabled = False
Game.Enabled = True
MoveNotes.Enabled = True
VScroll1.Enabled = False
End Sub

Private Sub mnuSaveSong_Click()
If Saved = False Then
    Save
    If Song.FileName <> "" Then
        'Open Song.FileName For Output As #1
        'Print #1, NumberofNotes
        'Close #1
    End If
    Saved = True
End If
End Sub

Private Sub AntiCheat_Timer()
SpaceCounter = 0
AntiCheat.Enabled = False
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
    Case 13
        AddBlue = False
        AddRed = False
        AddYellow = False
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

Private Sub Form_Terminate()
    Open App.Path & "\MissedNote.txt" For Output As #1
            Print #1, "exit"
            Close #1
End Sub

Private Sub Game_Timer()
CheckForCollision
For Index = 0 To NumberofNotes + 1
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

CheckForCollisionTimer.Enabled = True
Space = False
End Sub


Private Sub mnuSelectSong_Click()
'SelectSong.DefaultExt = ".wav"
'SelectSong.DialogTitle = "Select a song"
'SelectSong.Filter = "(*.wav)| *.wav"
'SelectSong.ShowOpen
'SongURL = SelectSong.FileName
If Saved = True Then
SongName = InputBox("Make sure the .wav is in the application folder and enter the name of the song (case sensitive): ")
Open Song.FileName & "\Song.txt" For Output As #1
Print #1, SongName
Close #1
Else
    Save
    SongName = InputBox("Make sure the .wav is in the application folder and enter the name of the song (case sensitive): ")
    Open Song.FileName & "\Song.txt" For Output As #1
    Print #1, SongName
    Close #1
End If
End Sub

Private Sub mnuStop_Click()
'Closes the sound engine
Open App.Path & "\MissedNote.txt" For Output As #1
Print #1, "exit"
Close #1
On Error GoTo ErrorHandler
Dim position As Integer
If Saved = True Then
    For BlueIndex = 0 To NumberofNotes Step 1
        BlueNote(BlueIndex).Visible = True
        Open Song.FileName & "\BlueNote" & BlueIndex & ".note" For Input As #1
        Input #1, position
        Close #1
        BlueNote(BlueIndex).Left = BlueLine
        BlueNote(BlueIndex).Top = position
    Next
    For RedIndex = 0 To NumberofNotes Step 1
        RedNote(RedIndex).Visible = True
        Open Song.FileName & "\RedNote" & RedIndex & ".note" For Input As #2
        Input #2, position
        Close #2
        RedNote(RedIndex).Left = RedLine
        RedNote(RedIndex).Top = position
    Next
    For YellowIndex = 0 To NumberofNotes Step 1
        YellowNote(YellowIndex).Visible = True
        Open Song.FileName & "\YellowNote" & YellowIndex & ".note" For Input As #3
        Input #3, position
        Close #3
        YellowNote(YellowIndex).Left = YellowLine
        YellowNote(YellowIndex).Top = position
    Next
End If
ErrorHandler:
Resume Next
shpRed.FillColor = &HFF&
shpGreen.FillColor = &H8000&
cmdAddBlue.Enabled = True
cmdAddYellow.Enabled = True
cmdAddRed.Enabled = True
fmeSpeed.Enabled = True
Game.Enabled = False
MoveNotes.Enabled = False
VScroll1.Enabled = True
End Sub

Private Sub mnuTrack1_Click()
AddBlue = True
BlueNote(BlueIndex).Top = imgBackGround.Top
BlueNote(BlueIndex).Visible = True
End Sub

Private Sub MoveNotes_Timer()
For Index = 0 To NumberofNotes
BlueNote(Index).Top = BlueNote(Index).Top + 50
RedNote(Index).Top = RedNote(Index).Top + 50
YellowNote(Index).Top = YellowNote(Index).Top + 50
Next
End Sub


Private Sub optBeginner_Click()
MoveNotes.Interval = Beginner
NoteSpeed = "Beginner"
End Sub

Private Sub optEasy_Click()
MoveNotes.Interval = Easy
NoteSpeed = "Easy"
End Sub

Private Sub optHard_Click()
MoveNotes.Interval = Hard
NoteSpeed = "Hard"
End Sub

Private Sub optImpossible_Click()
MoveNotes.Interval = Impossible
NoteSpeed = "Impossible"
End Sub

Private Sub optMedium_Click()
MoveNotes.Interval = Medium
NoteSpeed = "Medium"
End Sub

Private Sub tmrBlueExplosion_Timer()
    BaseBlueNote.Visible = False
    Select Case Counter1
    Case Is = 0
        imgBlueExplode(1).Visible = False
        imgBlueExplode(0).Visible = False
    Case Is = 1
        imgBlueExplode(1).Visible = True
        imgBlueExplode(0).Visible = False
    Case Is = 2
        imgBlueExplode(1).Visible = False
        imgBlueExplode(0).Visible = True
    Case Else
        BaseBlueNote.Visible = True
        imgBlueExplode(1).Visible = False
        imgBlueExplode(0).Visible = False
        tmrBlueExplosion.Enabled = False
    End Select
    Counter1 = Counter1 + 1
End Sub



Private Sub tmrRedExplosion_Timer()
    BaseRedNote.Visible = False
    Select Case Counter2
    Case Is = 0
        imgRedExplode(1).Visible = False
        imgRedExplode(0).Visible = False
    Case Is = 1
        imgRedExplode(1).Visible = True
        imgRedExplode(0).Visible = False
    Case Is = 2
        imgRedExplode(1).Visible = False
        imgRedExplode(0).Visible = True
    Case Else
        BaseRedNote.Visible = True
        Counter = 0
        imgRedExplode(1).Visible = False
        imgRedExplode(0).Visible = False
        tmrRedExplosion.Enabled = False
    End Select
    Counter2 = Counter2 + 1
End Sub

Private Sub tmrResetScrollbar_Timer()
'This is used to make sure the view on the track does not go beyond it's limits
If VScroll1.Value > 20200 Then
    VScroll1.Value = 20200
End If
Distance = VScroll1.Value
End Sub

Private Sub tmrYellowExplosion_Timer()
    Counter3 = Counter3 + 1
    BaseYellowNote.Visible = False
    Select Case Counter3
    Case Is = 0
        imgYellowExplode(1).Visible = False
        imgYellowExplode(0).Visible = False
    Case Is = 1
        imgYellowExplode(1).Visible = True
        imgYellowExplode(0).Visible = False
    Case Is = 2
        imgYellowExplode(1).Visible = False
        imgYellowExplode(0).Visible = True
    Case Else
        BaseYellowNote.Visible = True
        imgYellowExplode(1).Visible = False
        imgYellowExplode(0).Visible = False
        tmrYellowExplosion.Enabled = False
    End Select
    Counter3 = Counter3 + 1
End Sub

Private Sub VScroll1_Change()
If VScroll1.Value < 30000 Then
    imgBackGround.Top = -VScroll1.Value
Else
    imgBackGround.Top = 20200
End If
End Sub
