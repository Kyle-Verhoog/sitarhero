Attribute VB_Name = "CollisionFunctions"
Function Collision(Object1 As Object, Object2 As Object) As Boolean
Dim Ob1 As Object
Dim Ob2 As Object
Set Ob1 = Object1
Set Ob2 = Object2
If Ob1.Left < Ob2.Left + Ob2.Width And Ob1.Left + Ob1.Width > Ob2.Left And Ob1.Top < Ob2.Top + Ob2.Height And Ob1.Top + Ob1.Height > Ob2.Top Then
    Collision = True
Else
    Collision = False
End If
End Function
