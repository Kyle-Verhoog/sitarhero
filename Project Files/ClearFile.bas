Attribute VB_Name = "ClearFile"
Sub ClearCheerBoo()
On Error Resume Next:
            Open App.Path & "\Cheer-Boo.txt" For Output As #1
            Print #1, ""
            Close #1
End Sub
