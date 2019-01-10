VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   0  'None
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblTotalNotes 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   ":D"
      BeginProperty Font 
         Name            =   "Gill Sans MT Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1080
      TabIndex        =   3
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label lblVolume 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Full Volume"
      BeginProperty Font 
         Name            =   "Gill Sans MT Condensed"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label lblQuit 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Gill Sans MT Condensed"
         Size            =   33.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label lblDifficulty 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Easy"
      BeginProperty Font 
         Name            =   "Gill Sans MT Condensed"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Gill Sans MT Condensed"
         Size            =   33.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Easy As Integer = 25

Private Sub Form_Load()
Open App.Path & "\Difficulty.txt" For Input As #1
Input #1, SongName
Close #1

lblDifficulty.Caption = SongName
End Sub

Private Sub lblStats_Click()
End Sub

Private Sub lblDifficulty_Click()
If lblDifficulty.Caption = "easy" Then
    lblDifficulty.Caption = "medium"
    frmGame.MoveNotes.Interval = 15
ElseIf lblDifficulty.Caption = "medium" Then
    lblDifficulty.Caption = "hard"
    frmGame.MoveNotes.Interval = 12
ElseIf lblDifficulty.Caption = "hard" Then
    lblDifficulty.Caption = "impossible"
    frmGame.MoveNotes.Interval = 2
Else
    lblDifficulty.Caption = "easy"
    frmGame.MoveNotes.Interval = 20
End If
Open App.Path & "\Difficulty.txt" For Output As #1
Print #1, lblDifficulty.Caption;
Close #1
End Sub

Private Sub lblQuit_Click()
frmPauseMenu.Show
Unload Me
End Sub

Private Sub lblTotalNotes_Click()
'Change the traffic light's colour
'If frmGame.shpGreen.FillColor = &HFF00& Then
'    frmGame.shpGreen = &H8000&
'    'Light up Yellow
'    frmGame.shpYellow = &HFFFF&
'    frmGame.shpRed = &H80&
'ElseIf frmGame.shpYellow.FillColor = &HFFFF& Then
'    frmGame.shpGreen = &H8000&
'    frmGame.shpYellow = &H8080&
'    'Light up red
'    frmGame.shpRed = &HFF&
'ElseIf frmGame.shpRed = &HFF& Then
'    'Light up Green
'    frmGame.shpGreen = &HFF00&
'    frmGame.shpYellow = &H8080&
'    frmGame.shpRed = &H80&
'End If
End Sub

Private Sub lblVolume_Click()
WriteHitNote
End Sub
