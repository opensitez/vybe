

Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
        ' 1. Declare and Instantiate the class
        Dim user As New Person()

        ' 2. Set the property
        user.Name = "Alice"

        ' 3. Call the method
        user.Greet()
End Sub

