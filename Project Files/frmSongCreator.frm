VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSongCreator 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   12720
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12720
   ScaleWidth      =   9660
   Begin VB.PictureBox BaseRedNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   307
      Left            =   1920
      Picture         =   "frmSongCreator.frx":0000
      ScaleHeight     =   300
      ScaleWidth      =   630
      TabIndex        =   26
      Top             =   11280
      Width           =   630
   End
   Begin VB.PictureBox BaseYellowNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   307
      Left            =   3240
      Picture         =   "frmSongCreator.frx":0FF9
      ScaleHeight     =   300
      ScaleWidth      =   630
      TabIndex        =   25
      Top             =   11280
      Width           =   630
   End
   Begin VB.PictureBox BaseBlueNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   307
      Left            =   480
      Picture         =   "frmSongCreator.frx":201D
      ScaleHeight     =   300
      ScaleWidth      =   630
      TabIndex        =   24
      Top             =   11280
      Width           =   630
   End
   Begin VB.PictureBox YellowNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   307
      Index           =   2
      Left            =   6960
      Picture         =   "frmSongCreator.frx":2FF9
      ScaleHeight     =   300
      ScaleWidth      =   630
      TabIndex        =   23
      Top             =   6240
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox BlueNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   307
      Index           =   2
      Left            =   6960
      Picture         =   "frmSongCreator.frx":401D
      ScaleHeight     =   300
      ScaleWidth      =   630
      TabIndex        =   22
      Top             =   3720
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox RedNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   307
      Index           =   2
      Left            =   6960
      Picture         =   "frmSongCreator.frx":4FF9
      ScaleHeight     =   300
      ScaleWidth      =   630
      TabIndex        =   21
      Top             =   4800
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox YellowNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   307
      Index           =   1
      Left            =   6960
      Picture         =   "frmSongCreator.frx":5FF2
      ScaleHeight     =   300
      ScaleWidth      =   630
      TabIndex        =   20
      Top             =   6240
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox BlueNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   307
      Index           =   1
      Left            =   6960
      Picture         =   "frmSongCreator.frx":7016
      ScaleHeight     =   300
      ScaleWidth      =   630
      TabIndex        =   19
      Top             =   3720
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox RedNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   307
      Index           =   1
      Left            =   6960
      Picture         =   "frmSongCreator.frx":7FF2
      ScaleHeight     =   300
      ScaleWidth      =   630
      TabIndex        =   18
      Top             =   4800
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox RedNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   307
      Index           =   0
      Left            =   6960
      Picture         =   "frmSongCreator.frx":8FEB
      ScaleHeight     =   300
      ScaleWidth      =   630
      TabIndex        =   17
      Top             =   4800
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox BlueNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   307
      Index           =   0
      Left            =   6960
      Picture         =   "frmSongCreator.frx":9FE4
      ScaleHeight     =   300
      ScaleWidth      =   630
      TabIndex        =   16
      Top             =   3720
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox YellowNote 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   307
      Index           =   0
      Left            =   6960
      Picture         =   "frmSongCreator.frx":AFC0
      ScaleHeight     =   300
      ScaleWidth      =   630
      TabIndex        =   15
      Top             =   6240
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4095
      Left            =   5160
      Max             =   30000
      TabIndex        =   5
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton cmdAddYellow 
      Caption         =   "Add Yellow Note"
      Height          =   375
      Left            =   8040
      TabIndex        =   3
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddRed 
      Caption         =   "Add Red Note"
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddBlue 
      Caption         =   "Add Blue Note"
      Height          =   375
      Left            =   8040
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog Song 
      Left            =   8880
      Top             =   11400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrYellowExplosion 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   7440
      Top             =   9960
   End
   Begin VB.Timer tmrRedExplosion 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   7440
      Top             =   10680
   End
   Begin VB.Timer tmrSpecial 
      Interval        =   30
      Left            =   13200
      Top             =   11760
   End
   Begin VB.Timer tmrDisplayAnimation 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3720
      Top             =   6240
   End
   Begin VB.Timer CheckForCollisionTimer 
      Interval        =   700
      Left            =   3720
      Top             =   7200
   End
   Begin VB.Timer Game 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   7080
      Top             =   120
   End
   Begin VB.Timer MoveNotes 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   6120
      Tag             =   " "
      Top             =   3480
   End
   Begin VB.PictureBox imgBackGround 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   29400
      Left            =   -2160
      Picture         =   "frmSongCreator.frx":BFE4
      ScaleHeight     =   29400
      ScaleWidth      =   7215
      TabIndex        =   0
      Top             =   -16680
      Width           =   7215
      Begin VB.CommandButton cmdBottom 
         Enabled         =   0   'False
         Height          =   195
         Left            =   2520
         TabIndex        =   4
         Top             =   28800
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Timer tmrBlueExplosion 
         Enabled         =   0   'False
         Interval        =   70
         Left            =   6480
         Top             =   20520
      End
      Begin VB.Timer AntiCheat 
         Enabled         =   0   'False
         Interval        =   1200
         Left            =   6480
         Top             =   19560
      End
      Begin VB.PictureBox imgBlueExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   0
         Left            =   2640
         Picture         =   "frmSongCreator.frx":16F53
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   14
         Top             =   27960
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox imgRedExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   0
         Left            =   4080
         Picture         =   "frmSongCreator.frx":1A9E1
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   12
         Top             =   27960
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox imgYellowExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   307
         Index           =   0
         Left            =   5400
         Picture         =   "frmSongCreator.frx":1E46F
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   13
         Top             =   27960
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox imgBlueExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   900
         Index           =   2
         Left            =   2160
         Picture         =   "frmSongCreator.frx":21EFD
         ScaleHeight     =   900
         ScaleWidth      =   1890
         TabIndex        =   9
         Top             =   27720
         Visible         =   0   'False
         Width           =   1897
      End
      Begin VB.PictureBox imgRedExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   900
         Index           =   2
         Left            =   3480
         Picture         =   "frmSongCreator.frx":271E9
         ScaleHeight     =   900
         ScaleWidth      =   1890
         TabIndex        =   10
         Top             =   27720
         Visible         =   0   'False
         Width           =   1897
      End
      Begin VB.PictureBox imgYellowExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   900
         Index           =   2
         Left            =   4800
         Picture         =   "frmSongCreator.frx":2C4D5
         ScaleHeight     =   900
         ScaleWidth      =   1890
         TabIndex        =   11
         Top             =   27720
         Visible         =   0   'False
         Width           =   1897
      End
      Begin VB.PictureBox imgRedExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   900
         Index           =   1
         Left            =   3480
         Picture         =   "frmSongCreator.frx":317C1
         ScaleHeight     =   900
         ScaleWidth      =   1890
         TabIndex        =   6
         Top             =   27720
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.PictureBox imgBlueExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   900
         Index           =   1
         Left            =   2160
         Picture         =   "frmSongCreator.frx":3657E
         ScaleHeight     =   900
         ScaleWidth      =   1890
         TabIndex        =   7
         Top             =   27720
         Visible         =   0   'False
         Width           =   1897
      End
      Begin VB.PictureBox imgYellowExplode 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   900
         Index           =   1
         Left            =   4800
         Picture         =   "frmSongCreator.frx":3B33B
         ScaleHeight     =   900
         ScaleWidth      =   1890
         TabIndex        =   8
         Top             =   27720
         Visible         =   0   'False
         Width           =   1897
      End
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
      Begin VB.Menu mnuNewSong 
         Caption         =   "New Song"
      End
      Begin VB.Menu mnuOpenSong 
         Caption         =   "Open Song"
      End
      Begin VB.Menu mnuSaveSong 
         Caption         =   "Save Song"
      End
      Begin VB.Menu mnuSaveSongAs 
         Caption         =   "Save Song As"
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
Attribute VB_Name = "frmSongCreator"
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
Dim Repeat As Boolean
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
Const BlueLine As Integer = 465
Const RedLine As Integer = 1875
Const YellowLine As Integer = 3120





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
End Sub


Private Sub cmdAddBlue_Click()
AddBlue = True
BlueNote(BlueIndex).Visible = True
End Sub

Private Sub cmdAddRed_Click()
AddRed = True
RedNote(RedIndex).Visible = True
End Sub

Private Sub cmdAddYellow_Click()
AddYellow = True
YellowNote(YellowIndex).Visible = True
End Sub

Private Sub Form_Click()
If AddBlue = True Then
    Open App.Path & "\NoteFiles\BlueNote" & BlueIndex & ".note" For Output As #1
    Print #1, BlueNote(BlueIndex).Top
    Close #1
    BlueIndex = BlueIndex + 1
    AddBlue = False
End If

If AddRed = True Then
    Open App.Path & "\NoteFiles\RedNote" & RedIndex & ".note" For Output As #1
    Print #1, RedNote(RedIndex).Top
    Close #1
    RedIndex = RedIndex + 1
    AddRed = False
End If

If AddYellow = True Then
    Open App.Path & "\NoteFiles\YellowNote" & YellowIndex & ".note" For Output As #1
    Print #1, YellowNote(YellowIndex).Top
    Close #1
    YellowIndex = YellowIndex + 1
    NumberofNotes = NumberofNotes + 1
    AddYellow = False
End If
End Sub

Private Sub Form_Load()
'Always one less because of array (0) etc 10 notes would be 9
NumberofNotes = 0
BlueIndex = NumberofNotes
RedIndex = NumberofNotes
YellowIndex = NumberofNotes
HitNote = True
UpperBound = -600
LowerBound = -10
Score = 1000
Difficulty = 30
Multiplier = 10
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If AddBlue = True Then
    BlueNote(BlueIndex).Left = 530 '2700
    If Y < 9840 Then
        BlueNote(BlueIndex).Top = Y
    End If
 End If
 If AddRed = True Then
    RedNote(RedIndex).Left = 1920 '4080
    If Y < 9840 Then
        RedNote(RedIndex).Top = Y
    End If
 End If
 If AddYellow = True Then
    YellowNote(YellowIndex).Left = 3240 '5400
    If Y < 9840 Then
        YellowNote(YellowIndex).Top = Y
    End If
 End If
End Sub



Private Sub mnuOpenSong_Click()
    Song.DialogTitle = "Choose a song"
    Song.ShowOpen
    If Song.FileName = "" Then
        'User exited the dialog
    Else
        SongURL = Song.FileName
    End If
MsgBox SongURL
End Sub

Private Sub mnuPlay_Click()
shpGreen.FillColor = &HFF00&
shpRed.FillColor = &H80&
Shell App.Path & "\Launcher.exe"
cmdAddRed.Enabled = False
cmdAddBlue.Enabled = False
cmdAddYellow.Enabled = False
Game.Enabled = True
MoveNotes.Enabled = True
VScroll1.Enabled = False
End Sub

Private Sub mnuSaveSong_Click()
    Song.DefaultExt = ".song"
    Song.Filter = " (*.song)| *.Song"
    Song.DialogTitle = "Save your song"
    Song.ShowSave
    If Song.FileName <> "" Then
        Open Song.FileName For Output As #1
        Print #1, NumberofNotes
        Close #1
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

CheckForCollisionTimer.Enabled = True
Space = False
End Sub


Private Sub mnuStop_Click()
On Error GoTo ErrorHandler
Dim position As Integer
For BlueIndex = 0 To NumberofNotes Step 1
    BlueNote(BlueIndex).Visible = True
    Open App.Path & "\NoteFiles\BlueNote" & BlueIndex & ".note" For Input As #1
    Input #1, position
    Close #1
    BlueNote(BlueIndex).Left = BlueLine
    BlueNote(BlueIndex).Top = position
Next
For RedIndex = 0 To NumberofNotes Step 1
    RedNote(RedIndex).Visible = True
    Open App.Path & "\NoteFiles\RedNote" & RedIndex & ".note" For Input As #2
    Input #2, position
    Close #2
    RedNote(RedIndex).Left = RedLine
    RedNote(RedIndex).Top = position
Next
For YellowIndex = 0 To NumberofNotes Step 1
    YellowNote(YellowIndex).Visible = True
    Open App.Path & "\NoteFiles\YellowNote" & YellowIndex & ".note" For Input As #3
    Input #3, position
    Close #3
    YellowNote(YellowIndex).Left = YellowLine
    YellowNote(YellowIndex).Top = position
Next
ErrorHandler:

Resume Next
shpRed.FillColor = &HFF&
shpGreen.FillColor = &H8000&
Score = 1000
cmdAddBlue.Enabled = True
cmdAddYellow.Enabled = True
cmdAddRed.Enabled = True
Game.Enabled = False
MoveNotes.Enabled = False
VScroll1.Enabled = True
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

Private Sub VScroll1_Change()
imgBackGround.Top = -VScroll1.Value
End Sub
