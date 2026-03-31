Public Class KitchenSinkForm
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.TextBox1.Text = "Button Clicked!"
        Me.CheckBox1.Checked = False
        Me.Label1.Text = "Action Performed"
        Me.WebBrowser1.HTML = "<h2>Updated by Code</h2>"
    End Sub

    Private Sub KitchenSinkForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Verification logic could go here
    End Sub
End Class
