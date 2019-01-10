Attribute VB_Name = "MissedNote"
Sub MissedNoteFunction(ScoreAmount As Integer, Note As Object)
    WriteMissedNote
    Note.Visible = True
    Notes = 0
    Score = Score - ScoreAmount
    Missed = Missed + 1
    Score = Score - Difficulty
End Sub
