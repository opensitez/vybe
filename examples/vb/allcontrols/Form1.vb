Partial Class Form1

Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
    web1.Navigate("about:blank")
End Sub

Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
    txt1.Text = "hello"
    lbl1.Text = "Button clicked"
End Sub

Private Sub lst1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lst1.SelectedIndexChanged
    lbl1.Text = "List item " & lst1.SelectedIndex
End Sub

Private Sub cbo1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo1.SelectedIndexChanged
    lbl1.Text = "Color " & cbo1.SelectedIndex
End Sub

Private Sub tvw1_AfterSelect(sender As Object, e As EventArgs) Handles tvw1.AfterSelect
    lbl1.Text = "Node " & tvw1.SelectedNode.Text
End Sub

Private Sub toolButton_Click(sender As Object, e As EventArgs)
    lbl1.Text = "Toolbar clicked"
End Sub

Private Sub openMenuItem_Click(sender As Object, e As EventArgs)
    lbl1.Text = "Menu clicked"
End Sub
End Class
