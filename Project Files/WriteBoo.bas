Attribute VB_Name = "WriteBoo"
Sub Boo()
On Error Resume Next
            Open App.Path & "\Cheer-Boo.txt" For Output As #1
            Print #1, "boo"
            Close #1
End Sub

Sub EmptyFile()
On Error Resume Next
            Open App.Path & "\Cheer-Boo.txt" For Output As #1
            Print #1, ""
            Close #1
End Sub
