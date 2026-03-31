Partial Class Person
    Inherits System.Windows.Forms.Form

    Friend WithEvents btn1 As System.Windows.Forms.Button

    Private Sub InitializeComponent()
        Me.btn1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        Me.btn1.Location = New System.Drawing.Point(150, 180)
        Me.btn1.Size = New System.Drawing.Size(120, 30)
        Me.btn1.Text = "btn1"
        Me.btn1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btn1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btn1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btn1.Name = "btn1"
        Me.btn1.TabIndex = 0
        Me.Controls.Add(Me.btn1)
        Me.ClientSize = New System.Drawing.Size(640, 480)
        Me.Text = "Person"
        Me.Name = "Person"
        Me.ResumeLayout(False)
    End Sub
End Class
