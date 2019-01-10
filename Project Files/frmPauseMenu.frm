VERSION 5.00
Begin VB.Form frmPauseMenu 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblMenuButton 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Quit"
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
      Height          =   855
      Index           =   3
      Left            =   1080
      TabIndex        =   4
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label lblMenuButton 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Index           =   2
      Left            =   1080
      TabIndex        =   3
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label lblMenuButton 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Stats"
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
      Height          =   855
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label lblMenuButton 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Resume"
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
      Height          =   855
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Paused"
      BeginProperty Font 
         Name            =   "Gill Sans MT Condensed"
         Size            =   33.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   -120
      TabIndex        =   0
      Top             =   -120
      Width           =   4935
   End
End
Attribute VB_Name = "frmPauseMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Selected As Integer

Private Sub lblOptions_Click()
frmOptions.Show
End Sub

Private Sub lblQuit_Click()
Unload frmGame
Unload Me
End Sub

Private Sub lblResume_Click()
frmGame.Game.Enabled = True
frmGame.MoveNotes.Enabled = True
Unload Me
End Sub

Private Sub lblStats_Click()
frmStats.Show

End Sub

Private Sub lblMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMenuButton(Selected).BackColor = &H0&
End Sub

Private Sub lblMenuButton_Click(Index As Integer)
Select Case Index
Case 3
    Unload frmGame
    Unload frmStats
    Unload Me
Case 2
    frmPauseMenu.Hide
    frmOptions.Show
Case 1
    frmStats.Show
Case 0
    frmGame.Game.Enabled = True
    frmGame.MoveNotes.Enabled = True
    Unload Me
End Select
End Sub

Private Sub lblMenuButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Selected = Index
lblMenuButton(Index).BackColor = &HFF0000
End Sub
