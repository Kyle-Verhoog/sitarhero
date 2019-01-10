VERSION 5.00
Begin VB.Form frmStats 
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   ClientHeight    =   6975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8625
   BeginProperty Font 
      Name            =   "Gill Sans MT Condensed"
      Size            =   27.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrGraph 
      Interval        =   2000
      Left            =   1680
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Height          =   705
      Left            =   3960
      TabIndex        =   7
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   100
      Left            =   960
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      ScaleHeight     =   3015
      ScaleWidth      =   8295
      TabIndex        =   6
      Top             =   2880
      Width           =   8295
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "Performance: Score"
         Height          =   615
         Left            =   2640
         TabIndex        =   8
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.Label lblMenuButton 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Score Data"
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
      Index           =   5
      Left            =   360
      TabIndex        =   9
      Top             =   6120
      Width           =   3255
   End
   Begin VB.Shape Graph 
      BorderWidth     =   13
      Height          =   3255
      Left            =   0
      Top             =   2760
      Width           =   8535
   End
   Begin VB.Label lblMenuButton 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Return"
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
      Index           =   4
      Left            =   5160
      TabIndex        =   5
      Top             =   6120
      Width           =   3255
   End
   Begin VB.Label lblMenuButton 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Highscore:"
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
      Index           =   3
      Left            =   5040
      TabIndex        =   4
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label lblMenuButton 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "1/1 %100"
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
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label lblMenuButton 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Score: "
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
      Index           =   1
      Left            =   5040
      TabIndex        =   2
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label lblMenuButton 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Note Streak: "
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
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
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
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Selected As Integer
Dim Counter As Integer
Dim TempScore As Integer
Dim GraphScore(1 To 10) As Integer
Dim X1var As Integer
Dim X2var As Integer
Dim Y1var As Integer
Dim Y2var As Integer

Private Sub Form_Load()
X1var = 100
X2var = 1100
Y2var = 3000
Counter = 1
'Picture1.Line (100, TempScore)-(1100, 3000), vbRed, BF
Open App.Path & "\Highscore.txt" For Input As #1
Input #1, HighScore
Close #1
End Sub

Private Sub lblMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If lblMenuButton(5).Caption = "Score Data" Then
    lblMenuButton(Selected).BackColor = &H0&
    lblMenuButton(4).BackColor = &H0&
    Picture1.Line (0, GraphScore(1))-(900, 3000), vbRed, BF
    Picture1.Line (1000, GraphScore(2))-(1900, 3000), vbBlue, BF
    Picture1.Line (2000, GraphScore(3))-(2900, 3000), vbGreen, BF
    Picture1.Line (3000, GraphScore(4))-(3900, 3000), vbYellow, BF
    Picture1.Line (4000, GraphScore(5))-(4900, 3000), vbRed, BF
    Picture1.Line (5000, GraphScore(6))-(5900, 3000), vbBlue, BF
    Picture1.Line (6000, GraphScore(7))-(6900, 3000), vbGreen, BF
    Picture1.Line (7000, GraphScore(8))-(7900, 3000), vbYellow, BF
    Picture1.Line (8000, GraphScore(9))-(8900, 3000), vbRed, BF
'Else
'    Picture1.Line (0, GraphScore(1))-(900, 3000), vbRed, BF
'    Picture1.Line (1000, GraphScore(2))-(1900, 3000), vbBlue, BF
'    Picture1.Line (2000, GraphScore(3))-(2900, 3000), vbGreen, BF
'    Picture1.Line (3000, GraphScore(4))-(3900, 3000), vbYellow, BF
'    Picture1.Line (4000, GraphScore(5))-(4900, 3000), vbRed, BF
'    Picture1.Line (5000, GraphScore(6))-(5900, 3000), vbBlue, BF
'    Picture1.Line (6000, GraphScore(7))-(6900, 3000), vbGreen, BF
'    Picture1.Line (7000, GraphScore(8))-(7900, 3000), vbYellow, BF
'    Picture1.Line (8000, GraphScore(9))-(8900, 3000), vbRed, BF
'End If
End Sub

Private Sub lblMenuButton_Click(Index As Integer)
Select Case Index
Case 5
    If lblMenuButton(5).Caption = "Score Data" Then
        lblMenuButton(5).Caption = "Note Data"
        lbl1.Caption = "Performance: Notes"
    Else
        lblMenuButton(5).Caption = "Score Data"
        lbl1.Caption = "Performance: Score"
    End If
Case 4
    frmPauseMenu.Show
    Unload Me
End Select
End Sub

Private Sub lblMenuButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Selected = Index
lblMenuButton(Index).BackColor = &HFF0000
End Sub

Private Sub tmrGraph_Timer()
If Counter = 10 Then
    tmrGraph.Enabled = False
End If
'3350
TempScore = (TempScore - Score) + 4000
GraphScore(Counter) = TempScore
If Counter < 8 Then
'lblTitle(Counter).Caption = Score
End If
TempScore = 0
'X1var = X1var + 1100
'X2var = X2var + 1100
Counter = Counter + 1
End Sub

Private Sub tmrUpdate_Timer()
lblMenuButton(0).Caption = "Note Streak: " & Streak
lblMenuButton(1).Caption = "Current Score: " & Score
lblMenuButton(2).Caption = Hit & "/" & TotalNotes & " %" & Format(HitPercent, "Fixed")
lblMenuButton(3).Caption = "Highscore: " & HighScore
End Sub

Private Sub Command1_Click()
'(x1,y1),(x2,y1),(x1,y2) and (x2,y2)
Dim y1, y2, y3, y4 As Single
y1 = 100
y2 = 100
y3 = 100
y4 = 100
'2350 is the lowest
'MsgBox TempScore
'the first and third adjust the width
'the second and fourth adjust the height
Picture1.Line (100, 4000)-(1100, 3000), vbRed, BF
Picture1.Line (1200, 1800)-(2200, 3000), vbBlue, BF
Picture1.Line (2300, 1800)-(3300, 3000), vbGreen, BF
Picture1.Line (3400, 1800)-(4400, 3000), vbYellow, BF
Picture1.Line (4600, 1800)-(5500, 3000), vbOrange, BF
End Sub
