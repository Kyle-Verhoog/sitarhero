Attribute VB_Name = "WriteCheer"
Sub Cheer()
On Error Resume Next
            Open App.Path & "\Cheer-Boo.txt" For Output As #1
            Print #1, "cheer"
            Close #1
End Sub
