VERSION 5.00
Begin VB.Form frmGameOver 
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblStats 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Stats"
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
      Left            =   960
      TabIndex        =   4
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label lblQuit 
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
      Height          =   735
      Left            =   960
      TabIndex        =   2
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Score"
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
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label lblTotalNotes 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      Left            =   960
      TabIndex        =   0
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Game Over"
      BeginProperty Font 
         Name            =   "Gill Sans MT Condensed"
         Size            =   33.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmGameOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmGame.Game.Enabled = False
lblTotalNotes.Caption = Hit & "/" & TotalNotes & " %" & Format(HitPercent, "Fixed")
End Sub

Private Sub lblQuit_Click()
Unload frmPauseMenu
Unload frmGame
Unload Me
End Sub

Private Sub lblStats_Click()
frmStats.Show
End Sub

