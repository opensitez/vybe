Partial Class Form1
    Inherits System.Windows.Forms.Form

    Friend WithEvents btn0 As System.Windows.Forms.Button
    Friend WithEvents btn1 As System.Windows.Forms.Button
    Friend WithEvents btn2 As System.Windows.Forms.Button
    Friend WithEvents btn3 As System.Windows.Forms.Button
    Friend WithEvents btn4 As System.Windows.Forms.Button
    Friend WithEvents btn5 As System.Windows.Forms.Button
    Friend WithEvents btn6 As System.Windows.Forms.Button
    Friend WithEvents btn7 As System.Windows.Forms.Button
    Friend WithEvents btn8 As System.Windows.Forms.Button
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents btnReset As System.Windows.Forms.Button
    Friend WithEvents lbl12 As System.Windows.Forms.Label

    Private Sub InitializeComponent()
        Me.btn0 = New System.Windows.Forms.Button()
        Me.btn1 = New System.Windows.Forms.Button()
        Me.btn2 = New System.Windows.Forms.Button()
        Me.btn3 = New System.Windows.Forms.Button()
        Me.btn4 = New System.Windows.Forms.Button()
        Me.btn5 = New System.Windows.Forms.Button()
        Me.btn6 = New System.Windows.Forms.Button()
        Me.btn7 = New System.Windows.Forms.Button()
        Me.btn8 = New System.Windows.Forms.Button()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.btnReset = New System.Windows.Forms.Button()
        Me.lbl12 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        Me.btn0.Location = New System.Drawing.Point(20, 50)
        Me.btn0.Size = New System.Drawing.Size(60, 60)
        Me.btn0.Text = ""
        Me.btn0.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btn0.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btn0.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btn0.Name = "btn0"
        Me.btn0.TabIndex = 0
        Me.btn0.Enabled = True
        Me.btn0.Visible = True
        Me.btn1.Location = New System.Drawing.Point(80, 50)
        Me.btn1.Size = New System.Drawing.Size(60, 60)
        Me.btn1.Text = ""
        Me.btn1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btn1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btn1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btn1.Name = "btn1"
        Me.btn1.TabIndex = 1
        Me.btn1.Enabled = True
        Me.btn1.Visible = True
        Me.btn2.Location = New System.Drawing.Point(140, 50)
        Me.btn2.Size = New System.Drawing.Size(60, 60)
        Me.btn2.Text = ""
        Me.btn2.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btn2.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btn2.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btn2.Name = "btn2"
        Me.btn2.TabIndex = 2
        Me.btn2.Enabled = True
        Me.btn2.Visible = True
        Me.btn3.Location = New System.Drawing.Point(20, 110)
        Me.btn3.Size = New System.Drawing.Size(60, 60)
        Me.btn3.Text = ""
        Me.btn3.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btn3.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btn3.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btn3.Name = "btn3"
        Me.btn3.TabIndex = 3
        Me.btn3.Enabled = True
        Me.btn3.Visible = True
        Me.btn4.Location = New System.Drawing.Point(80, 110)
        Me.btn4.Size = New System.Drawing.Size(60, 60)
        Me.btn4.Text = ""
        Me.btn4.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btn4.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btn4.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btn4.Name = "btn4"
        Me.btn4.TabIndex = 4
        Me.btn4.Enabled = True
        Me.btn4.Visible = True
        Me.btn5.Location = New System.Drawing.Point(140, 110)
        Me.btn5.Size = New System.Drawing.Size(60, 60)
        Me.btn5.Text = ""
        Me.btn5.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btn5.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btn5.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btn5.Name = "btn5"
        Me.btn5.TabIndex = 5
        Me.btn5.Enabled = True
        Me.btn5.Visible = True
        Me.btn6.Location = New System.Drawing.Point(20, 170)
        Me.btn6.Size = New System.Drawing.Size(60, 60)
        Me.btn6.Text = ""
        Me.btn6.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btn6.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btn6.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btn6.Name = "btn6"
        Me.btn6.TabIndex = 6
        Me.btn6.Enabled = True
        Me.btn6.Visible = True
        Me.btn7.Location = New System.Drawing.Point(80, 170)
        Me.btn7.Size = New System.Drawing.Size(60, 60)
        Me.btn7.Text = ""
        Me.btn7.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btn7.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btn7.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btn7.Name = "btn7"
        Me.btn7.TabIndex = 7
        Me.btn7.Enabled = True
        Me.btn7.Visible = True
        Me.btn8.Location = New System.Drawing.Point(140, 170)
        Me.btn8.Size = New System.Drawing.Size(60, 60)
        Me.btn8.Text = ""
        Me.btn8.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btn8.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btn8.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btn8.Name = "btn8"
        Me.btn8.TabIndex = 8
        Me.btn8.Enabled = True
        Me.btn8.Visible = True
        Me.lblStatus.Location = New System.Drawing.Point(20, 230)
        Me.lblStatus.Size = New System.Drawing.Size(180, 23)
        Me.lblStatus.Text = "Player X Turn"
        Me.lblStatus.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.lblStatus.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.lblStatus.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.TabIndex = 9
        Me.lblStatus.Visible = True
        Me.btnReset.Location = New System.Drawing.Point(20, 260)
        Me.btnReset.Size = New System.Drawing.Size(180, 40)
        Me.btnReset.Text = "Reset Game"
        Me.btnReset.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btnReset.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btnReset.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.TabIndex = 10
        Me.btnReset.Enabled = True
        Me.btnReset.Visible = True
        Me.lbl12.Location = New System.Drawing.Point(32, 16)
        Me.lbl12.Size = New System.Drawing.Size(80, 20)
        Me.lbl12.Text = "lbl12"
        Me.lbl12.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.lbl12.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.lbl12.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.lbl12.Name = "lbl12"
        Me.lbl12.TabIndex = 0
        Me.lbl12.Visible = True
        Me.Controls.Add(Me.btn0)
        Me.Controls.Add(Me.btn1)
        Me.Controls.Add(Me.btn2)
        Me.Controls.Add(Me.btn3)
        Me.Controls.Add(Me.btn4)
        Me.Controls.Add(Me.btn5)
        Me.Controls.Add(Me.btn6)
        Me.Controls.Add(Me.btn7)
        Me.Controls.Add(Me.btn8)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.btnReset)
        Me.Controls.Add(Me.lbl12)
        AddHandler Me.Form1.Load, AddressOf Me.Form1_Load
        AddHandler Me.btn0.Click, AddressOf Me.btn0_Click
        AddHandler Me.btn1.Click, AddressOf Me.btn1_Click
        AddHandler Me.btn2.Click, AddressOf Me.btn2_Click
        AddHandler Me.btn3.Click, AddressOf Me.btn3_Click
        AddHandler Me.btn4.Click, AddressOf Me.btn4_Click
        AddHandler Me.btn5.Click, AddressOf Me.btn5_Click
        AddHandler Me.btn6.Click, AddressOf Me.btn6_Click
        AddHandler Me.btn7.Click, AddressOf Me.btn7_Click
        AddHandler Me.btn8.Click, AddressOf Me.btn8_Click
        AddHandler Me.btnReset.Click, AddressOf Me.btnReset_Click
        Me.ClientSize = New System.Drawing.Size(220, 320)
        Me.Text = "Tic Tac Toe"
        Me.Name = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub
End Class
