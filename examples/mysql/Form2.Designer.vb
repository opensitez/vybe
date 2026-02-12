Partial Class Form2
    Inherits System.Windows.Forms.Form

    Friend WithEvents da1 As System.Data.SqlClient.SqlDataAdapter
    Friend WithEvents bs1 As System.Windows.Forms.BindingSource
    Friend WithEvents txt1 As System.Windows.Forms.TextBox
    Friend WithEvents bnav1 As System.Windows.Forms.BindingNavigator

    Private Sub InitializeComponent()
        Me.da1 = New System.Data.SqlClient.SqlDataAdapter()
        Me.bs1 = New System.Windows.Forms.BindingSource()
        Me.txt1 = New System.Windows.Forms.TextBox()
        Me.bnav1 = New System.Windows.Forms.BindingNavigator()
        Me.SuspendLayout()
        Me.da1.ConnectionString = "Server=localhost;Port=3306;Database=genealogy;Uid=root;Pwd=password"
        Me.da1.Name = "da1"
        Me.bs1.DataSource = Me.da1
        Me.bs1.DataMember = "names"
        Me.bs1.Name = "bs1"
        Me.txt1.Location = New System.Drawing.Point(50, 100)
        Me.txt1.Size = New System.Drawing.Size(150, 25)
        Me.txt1.Text = ""
        Me.txt1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.txt1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.txt1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.txt1.Name = "txt1"
        Me.txt1.DataBindings.Add("Text", Me.bs1, "given")
        Me.txt1.TabIndex = 0
        Me.Controls.Add(Me.txt1)
        Me.bnav1.Location = New System.Drawing.Point(60, 400)
        Me.bnav1.Size = New System.Drawing.Size(300, 25)
        Me.bnav1.Text = "bnav1"
        Me.bnav1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.bnav1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.bnav1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.bnav1.Name = "bnav1"
        Me.bnav1.BindingSource = Me.bs1
        Me.bnav1.TabIndex = 0
        Me.Controls.Add(Me.bnav1)
        Me.ClientSize = New System.Drawing.Size(640, 480)
        Me.Text = "Form2"
        Me.Name = "Form2"
        Me.ResumeLayout(False)
    End Sub
End Class
