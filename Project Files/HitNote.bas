Attribute VB_Name = "HitNote"
Global SpaceCounter As Integer
Sub HitNoteFunction(ScoreAmount As Integer, NoteAmount As Integer, Note As Object)
    WriteHitNote
    Note.Visible = False
    Score = Score + ScoreAmount
    SpaceCounter = 0
    MissNote = False
    'Increasing the number of hit notes
    Hit = Hit + 1
    'If the score is less than 0 reset it to zero and then add on
    'ex if the score is at -5 and the person hits a note the score will be set to 1 instead of -4
    If Notes < 0 Then
        Notes = 0
        Notes = Notes + NoteAmount
    Else
        Notes = Notes + NoteAmount
    End If
End Sub
