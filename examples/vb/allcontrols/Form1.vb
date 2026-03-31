Partial Class Form1

Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
    web1.Navigate("https://example.com")
End Sub

Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
    txt1.Text = "hello"
End Sub
End Class
