Attribute VB_Name = "IO"
Option Explicit
Sub WriteHitNote()
On Error Resume Next:
            Open App.Path & "\MissedNote.txt" For Output As #1
            Print #1, "hit"
            Close #1
End Sub
Sub WriteMissedNote()
On Error Resume Next:
    Open App.Path & "\MissedNote.txt" For Output As #1
    Print #1, "missed"
    Close #1
End Sub
Function WriteFile(Contents As String, File As String) As String
On Error Resume Next:
    Open App.Path & File & ".txt" For Output As #1
    Print #1, Contents
    Close #1
End Function
Function ReadFile(Contents As String, File As String) As String
On Error Resume Next:
    Open App.Path & "\" & File & ".txt" For Input As #1
    Input #1, Contents
    Close #1
End Function
