Imports System.Windows.Forms

Partial Class Form1
    Inherits System.Windows.Forms.Form

    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox

    Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        Me.Button1.Location = New System.Drawing.Point(85, 90)
        Me.Button1.Size = New System.Drawing.Size(120, 30)
        Me.Button1.Text = "Click Me"
        Me.Button1.Name = "Button1"
        Me.Button1.TabIndex = 0
        Me.Label1.Location = New System.Drawing.Point(10, 10)
        Me.Label1.Size = New System.Drawing.Size(200, 23)
        Me.Label1.Text = "Hello World"
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 1
        Me.TextBox1.Location = New System.Drawing.Point(10, 40)
        Me.TextBox1.Size = New System.Drawing.Size(200, 23)
        Me.TextBox1.Text = ""
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.TabIndex = 2
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox1)
        Me.ClientSize = New System.Drawing.Size(640, 480)
        Me.Text = "Form1"
        Me.Name = "Form1"
        Me.ResumeLayout(False)
    End Sub
End Class
