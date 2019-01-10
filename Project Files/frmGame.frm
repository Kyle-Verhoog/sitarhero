VERSION 5.00
Begin VB.Form frmGame 
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   Caption         =   "Sitar Hero"
   ClientHeight    =   13500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   21600
   LinkTopic       =   "Form1"
   ScaleHeight     =   13500
   ScaleWidth      =   21600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer RealTime 
      Interval        =   1000
      Left            =   120
      Top             =   600
   End
   Begin VB.PictureBox BlueNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   382
      Index           =   3
      Left            =   4050
      Picture         =   "frmGame.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   915
      TabIndex        =   23
      Top             =   0
      Width           =   918
   End
   Begin VB.PictureBox RedNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   382
      Index           =   3
      Left            =   6240
      Picture         =   "frmGame.frx":32E6
      ScaleHeight     =   375
      ScaleWidth      =   915
      TabIndex        =   22
      Top             =   720
      Width           =   918
   End
   Begin VB.PictureBox YellowNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   382
      Index           =   3
      Left            =   8255
      Picture         =   "frmGame.frx":663F
      ScaleHeight     =   375
      ScaleWidth      =   930
      TabIndex        =   21
      Top             =   1200
      Width           =   937
   End
   Begin VB.PictureBox YellowNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   382
      Index           =   2
      Left            =   8255
      Picture         =   "frmGame.frx":992B
      ScaleHeight     =   375
      ScaleWidth      =   930
      TabIndex        =   20
      Top             =   3960
      Width           =   937
   End
   Begin VB.PictureBox BlueNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   382
      Index           =   2
      Left            =   4050
      Picture         =   "frmGame.frx":CC17
      ScaleHeight     =   375
      ScaleWidth      =   915
      TabIndex        =   19
      Top             =   3120
      Width           =   918
   End
   Begin VB.PictureBox RedNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   382
      Index           =   2
      Left            =   6240
      Picture         =   "frmGame.frx":FEFD
      ScaleHeight     =   375
      ScaleWidth      =   915
      TabIndex        =   18
      Top             =   3120
      Width           =   918
   End
   Begin VB.PictureBox RedNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   382
      Index           =   1
      Left            =   6240
      Picture         =   "frmGame.frx":13256
      ScaleHeight     =   375
      ScaleWidth      =   915
      TabIndex        =   17
      Top             =   7440
      Width           =   918
   End
   Begin VB.PictureBox BlueNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   382
      Index           =   1
      Left            =   4050
      Picture         =   "frmGame.frx":165AF
      ScaleHeight     =   375
      ScaleWidth      =   915
      TabIndex        =   16
      Top             =   6720
      Width           =   918
   End
   Begin VB.PictureBox YellowNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   382
      Index           =   0
      Left            =   8255
      Picture         =   "frmGame.frx":19895
      ScaleHeight     =   375
      ScaleWidth      =   930
      TabIndex        =   14
      Top             =   9480
      Width           =   937
   End
   Begin VB.PictureBox BlueNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   382
      Index           =   0
      Left            =   4050
      Picture         =   "frmGame.frx":1CB81
      ScaleHeight     =   375
      ScaleWidth      =   915
      TabIndex        =   13
      Top             =   9480
      Width           =   918
   End
   Begin VB.PictureBox RedNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   382
      Index           =   0
      Left            =   6240
      Picture         =   "frmGame.frx":1FE67
      ScaleHeight     =   375
      ScaleWidth      =   915
      TabIndex        =   12
      Top             =   8880
      Width           =   918
   End
   Begin VB.Timer tmrAnimation1 
      Interval        =   50
      Left            =   12600
      Top             =   840
   End
   Begin VB.CommandButton cmdBottom 
      Enabled         =   0   'False
      Height          =   435
      Left            =   3480
      TabIndex        =   11
      Top             =   12360
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Timer AntiCheat 
      Enabled         =   0   'False
      Interval        =   1200
      Left            =   10080
      Top             =   11040
   End
   Begin VB.Timer tmrBlueExplosion 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   4200
      Top             =   11760
   End
   Begin VB.Timer MoveNotes 
      Interval        =   20
      Left            =   9480
      Tag             =   " "
      Top             =   3600
   End
   Begin VB.Timer tmrRedExplosion 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   6480
      Top             =   11760
   End
   Begin VB.Timer tmrYellowExplosion 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   8520
      Top             =   11760
   End
   Begin VB.Timer tmrSpecial 
      Interval        =   30
      Left            =   2760
      Top             =   7800
   End
   Begin VB.Timer Game 
      Interval        =   30
      Left            =   1800
      Top             =   1080
   End
   Begin VB.PictureBox imgRedExplode 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   307
      Index           =   0
      Left            =   6360
      Picture         =   "frmGame.frx":231C0
      ScaleHeight     =   300
      ScaleWidth      =   630
      TabIndex        =   2
      Top             =   11040
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox imgYellowExplode 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   307
      Index           =   0
      Left            =   8400
      Picture         =   "frmGame.frx":26C4E
      ScaleHeight     =   300
      ScaleWidth      =   630
      TabIndex        =   3
      Top             =   11040
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox imgBlueExplode 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   307
      Index           =   0
      Left            =   4200
      Picture         =   "frmGame.frx":2A6DC
      ScaleHeight     =   300
      ScaleWidth      =   630
      TabIndex        =   4
      Top             =   11040
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox imgRedExplode 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   900
      Index           =   2
      Left            =   5760
      Picture         =   "frmGame.frx":2E16A
      ScaleHeight     =   900
      ScaleWidth      =   1890
      TabIndex        =   10
      Top             =   10800
      Visible         =   0   'False
      Width           =   1897
   End
   Begin VB.PictureBox imgYellowExplode 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   900
      Index           =   1
      Left            =   7800
      Picture         =   "frmGame.frx":33456
      ScaleHeight     =   900
      ScaleWidth      =   1890
      TabIndex        =   9
      Top             =   10800
      Visible         =   0   'False
      Width           =   1897
   End
   Begin VB.PictureBox imgRedExplode 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   900
      Index           =   1
      Left            =   5760
      Picture         =   "frmGame.frx":38213
      ScaleHeight     =   900
      ScaleWidth      =   1890
      TabIndex        =   7
      Top             =   10800
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.PictureBox imgYellowExplode 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   900
      Index           =   2
      Left            =   7800
      Picture         =   "frmGame.frx":3CFD0
      ScaleHeight     =   900
      ScaleWidth      =   1890
      TabIndex        =   5
      Top             =   10800
      Visible         =   0   'False
      Width           =   1897
   End
   Begin VB.PictureBox imgBlueExplode 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   900
      Index           =   1
      Left            =   3600
      Picture         =   "frmGame.frx":422BC
      ScaleHeight     =   900
      ScaleWidth      =   1890
      TabIndex        =   8
      Top             =   10800
      Visible         =   0   'False
      Width           =   1897
   End
   Begin VB.PictureBox imgBlueExplode 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   900
      Index           =   2
      Left            =   3600
      Picture         =   "frmGame.frx":47079
      ScaleHeight     =   900
      ScaleWidth      =   1890
      TabIndex        =   6
      Top             =   10800
      Visible         =   0   'False
      Width           =   1897
   End
   Begin VB.PictureBox YellowNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   382
      Index           =   1
      Left            =   8255
      Picture         =   "frmGame.frx":4C365
      ScaleHeight     =   375
      ScaleWidth      =   930
      TabIndex        =   15
      Top             =   6720
      Width           =   937
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Playing"
      BeginProperty Font 
         Name            =   "Gill Sans MT Condensed"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   600
      TabIndex        =   30
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label lblUsername 
      BackStyle       =   0  'Transparent
      Caption         =   "Username: "
      BeginProperty Font 
         Name            =   "Gill Sans MT Condensed"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   29
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label lblRealTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Gill Sans MT Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   28
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label lblSongName 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Song: "
      BeginProperty Font 
         Name            =   "Gill Sans MT Condensed"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   27
      Top             =   6840
      Width           =   3975
   End
   Begin VB.Label lblScoreDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "ConcursoItalian BTN"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   14400
      TabIndex        =   0
      Top             =   9240
      Width           =   3495
   End
   Begin VB.Label lblYellowPlus10 
      BackStyle       =   0  'Transparent
      Caption         =   "+10"
      BeginProperty Font 
         Name            =   "Gill Sans MT Condensed"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   9120
      TabIndex        =   26
      Top             =   11880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblRedPlus10 
      BackStyle       =   0  'Transparent
      Caption         =   "+10"
      BeginProperty Font 
         Name            =   "Gill Sans MT Condensed"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   7080
      TabIndex        =   25
      Top             =   11880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblBluePlus10 
      BackStyle       =   0  'Transparent
      Caption         =   "+10"
      BeginProperty Font 
         Name            =   "Gill Sans MT Condensed"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4800
      TabIndex        =   24
      Top             =   11880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image BaseYellowNote 
      Height          =   375
      Left            =   8255
      Picture         =   "frmGame.frx":4F651
      Top             =   11040
      Width           =   930
   End
   Begin VB.Image BaseRedNote 
      Height          =   390
      Left            =   6240
      Picture         =   "frmGame.frx":5293D
      Top             =   11040
      Width           =   915
   End
   Begin VB.Image BaseBlueNote 
      Height          =   375
      Left            =   4050
      Picture         =   "frmGame.frx":55C96
      Top             =   11040
      Width           =   915
   End
   Begin VB.Image img1 
      Height          =   3585
      Left            =   14040
      Picture         =   "frmGame.frx":58F7C
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Image img2 
      Height          =   3585
      Left            =   14040
      Picture         =   "frmGame.frx":5C68D
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Image img3 
      Height          =   3585
      Left            =   14040
      Picture         =   "frmGame.frx":5CA76
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label lblStreakDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "ConcursoItalian BTN"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Shape shpGreen 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   12120
      Shape           =   3  'Circle
      Top             =   11880
      Width           =   1335
   End
   Begin VB.Shape shpRed 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   12120
      Shape           =   3  'Circle
      Top             =   9240
      Width           =   1335
   End
   Begin VB.Shape shpYellow 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008080&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   12120
      Shape           =   3  'Circle
      Top             =   10560
      Width           =   1335
   End
   Begin VB.Image SelectBlueNote 
      Height          =   375
      Left            =   4050
      Picture         =   "frmGame.frx":5CE5F
      Top             =   11040
      Width           =   915
   End
   Begin VB.Image SelectRedNote 
      Height          =   390
      Left            =   6240
      Picture         =   "frmGame.frx":605D4
      Top             =   11040
      Width           =   915
   End
   Begin VB.Image SelectYellowNote 
      Height          =   375
      Left            =   8250
      Picture         =   "frmGame.frx":63D87
      Top             =   11040
      Width           =   930
   End
   Begin VB.Image imgNeutral 
      Enabled         =   0   'False
      Height          =   13500
      Left            =   0
      Picture         =   "frmGame.frx":6740C
      Top             =   0
      Width           =   21600
   End
   Begin VB.Image imgFail 
      Enabled         =   0   'False
      Height          =   13500
      Left            =   0
      Picture         =   "frmGame.frx":7D77C
      Top             =   0
      Visible         =   0   'False
      Width           =   21600
   End
   Begin VB.Image imgWinning 
      Enabled         =   0   'False
      Height          =   13500
      Left            =   0
      Picture         =   "frmGame.frx":96FA5
      Top             =   0
      Visible         =   0   'False
      Width           =   21600
   End
   Begin VB.Image imgHalfWin 
      Enabled         =   0   'False
      Height          =   13500
      Left            =   0
      Picture         =   "frmGame.frx":B0E3A
      Top             =   0
      Visible         =   0   'False
      Width           =   21600
   End
   Begin VB.Image imgHalfFail 
      Enabled         =   0   'False
      Height          =   13500
      Left            =   0
      Picture         =   "frmGame.frx":CB6E9
      Top             =   0
      Visible         =   0   'False
      Width           =   21600
   End
End
Attribute VB_Name = "frmGame"
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
Dim SongTime As Long
Dim SongName As String
Dim Time As Long



Private Sub BottomCollision()
For Index = 0 To NumberofNotes
    If BlueNote(Index).Visible = False Then
        Randomize
        RandomPosition = Int(Rnd * UpperBound) + LowerBound
        BlueNote(Index).Top = RandomPosition
        BlueNote(Index).Visible = True
    End If
    
    If RedNote(Index).Visible = False Then
        Randomize
        RandomPosition = Int(Rnd * UpperBound) + LowerBound
        RedNote(Index).Top = RandomPosition
        RedNote(Index).Visible = True
    End If
    
    If YellowNote(Index).Visible = False Then
        Randomize
        RandomPosition = Int(Rnd * UpperBound) + LowerBound
        YellowNote(Index).Top = RandomPosition
        YellowNote(Index).Visible = True
    End If
    
    'These lines of code are the base in which notes hit and count as a miss
    If BlueNote(Index).Visible = True Then
        If Collision(BlueNote(Index), cmdBottom) = True Then
            WriteMissedNote
            Score = Score - Difficulty
            'Setting the note to a slightly different height and bumping it back up
            Randomize
            RandomPosition = Int(Rnd * UpperBound) + LowerBound
            BlueNote(Index).Top = RandomPosition
            BlueNote(Index).Visible = True
            MissNote = True
            HitNote = False
            Missed = Missed + 1
            'Sets the streak to 0 if it is greater else it will keep subtracting from it
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
            Randomize
            RandomPosition = Int(Rnd * UpperBound) + LowerBound
            RedNote(Index).Top = RandomPosition
            RedNote(Index).Visible = True
            MissNote = True
            HitNote = False
            Missed = Missed + 1
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
            Randomize
            RandomPosition = Int(Rnd * UpperBound) + LowerBound
            YellowNote(Index).Top = RandomPosition
            YellowNote(Index).Visible = True
            MissNote = True
            HitNote = False
            Missed = Missed + 1
            If Notes > 0 Then
                Notes = 0
            Else
                Notes = Notes - 1
            End If
        End If
    End If
Next
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
        SelectBlueNote.Visible = True
        BaseBlueNote.Visible = False
    Case 113
        F2 = True
        SelectRedNote.Visible = True
        BaseRedNote.Visible = False
    Case 114
        F3 = True
        SelectYellowNote.Visible = True
        BaseYellowNote.Visible = False
    Case 49
        F1 = True
        SelectBlueNote.Visible = True
        BaseBlueNote.Visible = False
    Case 50
        F2 = True
        SelectRedNote.Visible = True
        BaseRedNote.Visible = False
    Case 51
        F3 = True
        SelectYellowNote.Visible = True
        BaseYellowNote.Visible = False
    Case 40
        Space = True
        SpaceCounter = SpaceCounter + 1
        AntiCheat.Enabled = True
    Case 32
        'tmrAnimation1.Enabled = True
        Space = True
        SpaceCounter = SpaceCounter + 1
        AntiCheat.Enabled = True
    Case 13
        Space = True
        SpaceCounter = SpaceCounter + 1
        AntiCheat.Enabled = True
    Case 8
        Space = True
        SpaceCounter = SpaceCounter + 1
        AntiCheat.Enabled = True
    Case 76
        Unload frmStats
        Unload Me
    Case 27
        frmPauseMenu.Show
        MoveNotes.Enabled = False
        'lblMenu.Visible = True
        'lblQuit.Visible = True
        'lblStats.Visible = True
        'lblResume.Visible = True
    Case 80
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
tmrAnimation1.Enabled = False
Select Case KeyCode
    Case 112
        F1 = False
        SelectBlueNote.Visible = False
        BaseBlueNote.Visible = True
    Case 113
        F2 = False
        SelectRedNote.Visible = False
        BaseRedNote.Visible = True
    Case 114
        F3 = False
        SelectYellowNote.Visible = False
        BaseYellowNote.Visible = True
    Case 115
        F4 = False
    Case 49
        F1 = False
        SelectBlueNote.Visible = False
        BaseBlueNote.Visible = True
    Case 50
        F2 = False
        SelectRedNote.Visible = False
        BaseRedNote.Visible = True
    Case 51
        F3 = False
        SelectYellowNote.Visible = False
        BaseYellowNote.Visible = True
    Case 40
        Space = False
    Case 32
        Space = False
End Select
End Sub

Private Sub Form_Load()
Open App.Path & "\Username.txt" For Input As #1
Input #1, Username
Close #1

lblUsername.Caption = "User: " & Username

frmStats.Show
frmStats.Hide
Open App.Path & "\Song.txt" For Input As #1
Input #1, SongName
Close #1
lblSongName.Caption = lblSongName.Caption & " " & SongName

On Error Resume Next:
'Check the difficulty and ajust the speed accordingly
Open App.Path & "\Difficulty.txt" For Input As #1
Input #1, FileContents
Close #1

'Adjust the difficulty according to the file
If FileContents = "easy" Then
    MoveNotes.Interval = 25
ElseIf FileContents = "medium" Then
    MoveNotes.Interval = 15
ElseIf FileContents = "hard" Then
    MoveNotes.Interval = 12
ElseIf FileContents = "impossible" Then
    MoveNotes.Interval = 2
Else
    MoveNotes.Interval = 20
End If


'Get the song time from the file
Open App.Path & "\Length.txt" For Input As #1
Input #1, SongTime
Close #1

imgNeutral.Picture = LoadPicture(App.Path & "\Images\neutral.jpg")
imgHalfWin.Picture = LoadPicture(App.Path & "\Images\halfwaywinning.jpg")
imgHalfFail.Picture = LoadPicture(App.Path & "\Images\halfwayfailing.jpg")
imgWinning.Picture = LoadPicture(App.Path & "\Images\winning.jpg")
imgFail.Picture = LoadPicture(App.Path & "\Images\failing.jpg")
Open App.Path & "\MissedNote.txt" For Output As #1
Print #1, "hit"
Close #1

Open App.Path & "\Command.txt" For Output As #1
Print #1, "start";
Close #1

EmptyFile
'Always one less because of array (0)
NumberofNotes = 3
HitNote = True
Score = 1000
Difficulty = 30
Multiplier = 10
UpperBound = -100
LowerBound = -600
Duration.Enabled = True
End Sub

Private Sub Form_Terminate()
On Error Resume Next
    Open App.Path & "\MissedNote.txt" For Output As #1
            Print #1, "exit"
            Close #1
End Sub
Private Sub StreakMultiplier()
If Notes > Streak Then
    Streak = Notes
End If
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
        lblBluePlus10.Caption = "+20"
        lblRedPlus10.Caption = "+20"
        lblYellowPlus10.Caption = "+20"
    Case Is > 20
        Multiplier = 15
        lblBluePlus10.Caption = "+15"
        lblRedPlus10.Caption = "+15"
        lblYellowPlus10.Caption = "+15"
        'frmGame.BackColor = &HFF00&
    Case Else
        lblBluePlus10.Caption = "+10"
        lblRedPlus10.Caption = "+10"
        lblYellowPlus10.Caption = "+10"
        frmGame.BackColor = &H80FF&
        Difficulty = 30
        Multiplier = 10
End Select

Select Case Notes
    Case Is >= 20
        lblStreakDisplay.ForeColor = &HFF00&
    Case Is >= 10
        lblStreakDisplay.ForeColor = &HFFFF&
    Case Else
        lblStreakDisplay.ForeColor = &HFF&
End Select

End Sub
Private Sub NoteCollision()
'Detect collisions with the notes on the base notes as well as with the proper button(s) clicked
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
End Sub
Private Sub ScoreActions()
Select Case Score
    Case Is <= 500
        'Make the crowd boo
        Boo
        imgFail.Visible = True
        imgHalfFail.Visible = False
        shpRed.FillColor = &HFF&
        shpGreen.FillColor = &H8000&
        shpYellow.FillColor = &H8080&
    Case Is <= 700
        'Change cheer-boo file back so that repeating is not an issue
        EmptyFile
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
        imgHalfWin.Visible = True
        imgNeutral.Visible = False
        imgWinning.Visible = False
        shpGreen.FillColor = &HFF00&
        shpRed.FillColor = &H80&
        shpYellow.FillColor = &H8080&
    Case Else
        EmptyFile
        imgNeutral.Visible = True
        imgHalfWin.Visible = False
        imgWinning.Visible = False
        imgHalfFail.Visible = False
        shpYellow.FillColor = &HFFFF&
        shpRed.FillColor = &H80&
        shpGreen.FillColor = &H8000&
End Select
End Sub
Private Sub Game_Timer()
lblRealTime.Caption = Time
BottomCollision
StreakMultiplier
NoteCollision
ScoreActions
'Make sure the select note object is visible when selected
If F1 = True Then
    SelectBlueNote.Visible = True
    BaseBlueNote.Visible = False
Else
    SelectBlueNote.Visible = False
    BaseBlueNote.Visible = True
End If

If F2 = True Then
    SelectRedNote.Visible = True
    BaseRedNote.Visible = False
Else
    SelectRedNote.Visible = False
    BaseRedNote.Visible = True
End If

If F3 = True Then
    SelectYellowNote.Visible = True
    BaseYellowNote.Visible = False
Else
    SelectYellowNote.Visible = False
    BaseYellowNote.Visible = True
End If

On Error Resume Next:
'Reading the file to see if the music is playing full volume to play the animation
Open App.Path & "\MissedNote.txt" For Input As #1
Input #1, FileContents
Close #1

If FileContents = "hit" Then
    tmrAnimation1.Enabled = True
Else
    tmrAnimation1.Enabled = False
End If
If Score <= 0 Then
    On Error Resume Next:
    'Tell the sound engine to make the crowd boo
    Boo
    'Tell the Sound Engine to stop playing music
    Open App.Path & "\MissedNote.txt" For Output As #1
        Print #1, "exit"
    Close #1
    
    Game.Enabled = False
        'Pop up the menu
        frmGameOver.Show
End If
TotalNotes = Hit + Missed
If Hit > 0 Then
    HitPercent = (Hit / TotalNotes) * 100
End If

If Missed > 0 Then
    MissPercent = (Missed / TotalNotes) * 100
End If
lblStreakDisplay.Caption = Notes
Space = False
lblScoreDisplay.Caption = Score

'If the times meet the song is over
If Time = SongTime Then
    frmGameOver.Show
End If
End Sub

Private Sub lblQuit_Click()
Unload Me
End Sub

Private Sub lblResume_Click()
MoveNotes.Enabled = True
Game.Enabled = True
lblMenu.Visible = False
lblQuit.Visible = False
lblStats.Visible = False
lblResume.Visible = False
End Sub



Private Sub MoveNotes_Timer()
For Index = 0 To NumberofNotes
    BlueNote(Index).Top = BlueNote(Index).Top + 50
    RedNote(Index).Top = RedNote(Index).Top + 50
    YellowNote(Index).Top = YellowNote(Index).Top + 50
Next
End Sub

Private Sub RealTime_Timer()
'Timer does not support larger than integer data :@ not good for putting the number of seconds in a file :|
'So we shall improvise
'Time is a long :)
Time = Time + 1
End Sub

Private Sub tmrAnimation1_Timer()
If img3.Visible = True Then
    img3.Visible = False
    img2.Visible = True
ElseIf img2.Visible = True Then
    img2.Visible = False
    img1.Visible = True
ElseIf img1.Visible = True Then
    img1.Visible = False
    img3.Visible = True
End If
End Sub

Private Sub tmrBlueExplosion_Timer()
    If lblBluePlus10.Visible = True Then
        lblBluePlus10.Visible = False
    Else
        lblBluePlus10.Visible = True
    End If
    
    BaseBlueNote.Visible = False
    Select Case Counter1
    Case Is = 0
        imgBlueExplode(0).Visible = True
        imgBlueExplode(1).Visible = False
        imgBlueExplode(2).Visible = False
    Case Is = 1
        imgBlueExplode(0).Visible = False
        imgBlueExplode(1).Visible = True
        imgBlueExplode(2).Visible = False
    Case Is = 2
        imgBlueExplode(0).Visible = False
        imgBlueExplode(1).Visible = False
        imgBlueExplode(2).Visible = True
    Case Else
        BaseBlueNote.Visible = True
        imgBlueExplode(0).Visible = False
        imgBlueExplode(1).Visible = False
        imgBlueExplode(2).Visible = False
        tmrBlueExplosion.Enabled = False
    End Select
    Counter1 = Counter1 + 1
End Sub

Private Sub tmrRedExplosion_Timer()
    If lblRedPlus10.Visible = True Then
        lblRedPlus10.Visible = False
    Else
        lblRedPlus10.Visible = True
    End If
    
    BaseRedNote.Visible = False
    Select Case Counter2
    Case Is = 0
        imgRedExplode(0).Visible = True
        imgRedExplode(1).Visible = False
        imgRedExplode(2).Visible = False
    Case Is = 1
        imgRedExplode(0).Visible = False
        imgRedExplode(1).Visible = True
        imgRedExplode(2).Visible = False
    Case Is = 2
        imgRedExplode(0).Visible = False
        imgRedExplode(1).Visible = False
        imgRedExplode(2).Visible = True
    Case Else
        BaseRedNote.Visible = True
        Counter = 0
        imgRedExplode(0).Visible = False
        imgRedExplode(1).Visible = False
        imgRedExplode(2).Visible = False
        tmrRedExplosion.Enabled = False
    End Select
    Counter2 = Counter2 + 1
End Sub

Private Sub tmrYellowExplosion_Timer()
    If lblYellowPlus10.Visible = True Then
        lblYellowPlus10.Visible = False
    Else
        lblYellowPlus10.Visible = True
    End If
    
    BaseYellowNote.Visible = False
    Select Case Counter3
    Case Is = 0
        imgYellowExplode(0).Visible = True
        imgYellowExplode(1).Visible = False
        imgYellowExplode(2).Visible = False
    Case Is = 1
        imgYellowExplode(0).Visible = False
        imgYellowExplode(1).Visible = True
        imgYellowExplode(2).Visible = False
    Case Is = 2
        imgYellowExplode(0).Visible = False
        imgYellowExplode(1).Visible = False
        imgYellowExplode(2).Visible = True
    Case Else
        BaseYellowNote.Visible = True
        imgYellowExplode(0).Visible = False
        imgYellowExplode(1).Visible = False
        imgYellowExplode(2).Visible = False
        tmrYellowExplosion.Enabled = False
    End Select
    Counter3 = Counter3 + 1
End Sub
