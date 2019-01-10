Attribute VB_Name = "SaveProject"
Sub Save()
    frmSongEditor.Song.DefaultExt = ""
    frmSongEditor.Song.Filter = " (Folder)| *"
    frmSongEditor.Song.DialogTitle = "Save your song"
    frmSongEditor.Song.ShowSave
    Saved = True
    On Error Resume Next:
            MkDir frmSongEditor.Song.FileName & "\"
End Sub

