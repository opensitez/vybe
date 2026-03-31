Partial Class KitchenSinkForm
    Inherits System.Windows.Forms.Form

    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents WebBrowser1 As System.Windows.Forms.WebBrowser
    Friend WithEvents Button1 As System.Windows.Forms.Button

    Private Sub InitializeComponent()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.WebBrowser1 = New System.Windows.Forms.WebBrowser()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        Me.TextBox1.Location = New System.Drawing.Point(10, 50)
        Me.TextBox1.Size = New System.Drawing.Size(200, 100)
        Me.TextBox1.Text = "Initial Text"
        Me.TextBox1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.TextBox1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.TextBox1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.TabIndex = 0
        Me.TextBox1.Enabled = True
        Me.TextBox1.Multiline = True
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Visible = True
        Me.CheckBox1.Location = New System.Drawing.Point(260, 130)
        Me.CheckBox1.Size = New System.Drawing.Size(100, 24)
        Me.CheckBox1.Text = "Check Me"
        Me.CheckBox1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.CheckBox1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.CheckBox1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.TabIndex = 1
        Me.CheckBox1.Checked = True
        Me.CheckBox1.Enabled = True
        Me.CheckBox1.Value = 0
        Me.CheckBox1.Visible = True
        Me.Label1.Location = New System.Drawing.Point(20, 170)
        Me.Label1.Size = New System.Drawing.Size(100, 23)
        Me.Label1.Text = "Label Text"
        Me.Label1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.Label1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 2
        Me.Label1.Visible = True
        Me.WebBrowser1.Location = New System.Drawing.Point(10, 220)
        Me.WebBrowser1.Size = New System.Drawing.Size(300, 200)
        Me.WebBrowser1.Text = "WebBrowser1"
        Me.WebBrowser1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.WebBrowser1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.WebBrowser1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.WebBrowser1.Name = "WebBrowser1"
        Me.WebBrowser1.TabIndex = 3
        Me.WebBrowser1.Enabled = True
        Me.WebBrowser1.HTML = "<h1>Kitchen Sink</h1><p>Test Content</p>"
        Me.WebBrowser1.URL = "about:blank"
        Me.WebBrowser1.Visible = True
        Me.Button1.Location = New System.Drawing.Point(260, 80)
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.Text = "Click Me"
        Me.Button1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.Button1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.Button1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.Button1.Name = "Button1"
        Me.Button1.TabIndex = 4
        Me.Button1.Enabled = True
        Me.Button1.Visible = True
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.WebBrowser1)
        Me.Controls.Add(Me.Button1)
        AddHandler Me.Button1.Click, AddressOf Me.Button1_Click
        AddHandler Me.KitchenSinkForm.Load, AddressOf Me.KitchenSinkForm_Load
        Me.ClientSize = New System.Drawing.Size(400, 450)
        Me.Text = "Kitchen Sink Test"
        Me.Name = "KitchenSinkForm"
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub
End Class
