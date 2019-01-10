Attribute VB_Name = "Variables"
'GAME VARIABLES---------------------
'Used for keeping score
Global Score As Integer

Global Username As String

'Used for number of hit or missed notes
Global Missed As Single
Global Hit As Single
Global TotalNotes As Single
Global HitPercent As Single
Global MissPercent As Single

'Used for checking whether space has been pressed
Global Space As Boolean
'Used for checking whether the number keys or function keys have been pressed
Global F1, F2, F3, F4 As Boolean

'Sets the amount lost per note missed
Global Difficulty As Integer
'Sets the amount gained per note hit
Global Multiplier As Integer

'Keeps count of the streak
Global Notes As Integer
Global Streak As Integer
'Used for keeping highscore
Global HighScore As Integer

'Used for the upper and lower bounds of random functions
Global UpperBound, LowerBound As Integer

'Used to store a random position for the note
Global RandomPosition As Integer

Global FileContents As String

Global HitNote As Boolean
Global MissNote As Boolean

Global Finished As String
Global Index As Integer
Global Index2 As Integer
Global Index3 As Integer
Global Collide  As Boolean
Global NumberofNotes As Integer


Global SongTitle As String
Global SongLength As Integer
'Editor Variables----------------------
Global SongURL As String

