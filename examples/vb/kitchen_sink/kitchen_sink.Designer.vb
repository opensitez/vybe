Public Class KitchenSinkForm
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

        ' TextBox1
        Me.TextBox1.Location = New System.Drawing.Point(10, 10)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(200, 100)
        Me.TextBox1.TabIndex = 0
        Me.TextBox1.Text = "Initial Text"
        Me.TextBox1.Multiline = True
        Me.TextBox1.ReadOnly = True

        ' CheckBox1
        Me.CheckBox1.Location = New System.Drawing.Point(10, 120)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(100, 24)
        Me.CheckBox1.TabIndex = 1
        Me.CheckBox1.Text = "Check Me"
        Me.CheckBox1.Checked = True

        ' Label1
        Me.Label1.Location = New System.Drawing.Point(10, 150)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 23)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Label Text"

        ' WebBrowser1
        Me.WebBrowser1.Location = New System.Drawing.Point(10, 180)
        Me.WebBrowser1.Name = "WebBrowser1"
        Me.WebBrowser1.Size = New System.Drawing.Size(300, 200)
        Me.WebBrowser1.TabIndex = 3
        Me.WebBrowser1.HTML = "<h1>Kitchen Sink</h1><p>Test Content</p>"

        ' Button1
        Me.Button1.Location = New System.Drawing.Point(10, 390)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Click Me"

        ' KitchenSinkForm
        Me.ClientSize = New System.Drawing.Size(400, 450)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.WebBrowser1)
        Me.Controls.Add(Me.Button1)
        Me.Name = "KitchenSinkForm"
        Me.Text = "Kitchen Sink Test"
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub
End Class
