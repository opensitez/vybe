Partial Class Form1
    Inherits System.Windows.Forms.Form

    Friend WithEvents btnOpen As System.Windows.Forms.Button
    Friend WithEvents btnColor As System.Windows.Forms.Button
    Friend WithEvents btnMsgbox As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnFont As System.Windows.Forms.Button
    Friend WithEvents btnInput As System.Windows.Forms.Button

    Private Sub InitializeComponent()
        Me.btnOpen = New System.Windows.Forms.Button()
        Me.btnColor = New System.Windows.Forms.Button()
        Me.btnMsgbox = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnFont = New System.Windows.Forms.Button()
        Me.btnInput = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        Me.btnOpen.Location = New System.Drawing.Point(100, 110)
        Me.btnOpen.Size = New System.Drawing.Size(120, 30)
        Me.btnOpen.Text = "Open"
        Me.btnOpen.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btnOpen.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btnOpen.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.TabIndex = 0
        Me.Controls.Add(Me.btnOpen)
        Me.btnColor.Location = New System.Drawing.Point(100, 170)
        Me.btnColor.Size = New System.Drawing.Size(120, 30)
        Me.btnColor.Text = "Color"
        Me.btnColor.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btnColor.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btnColor.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btnColor.Name = "btnColor"
        Me.btnColor.TabIndex = 0
        Me.Controls.Add(Me.btnColor)
        Me.btnMsgbox.Location = New System.Drawing.Point(100, 220)
        Me.btnMsgbox.Size = New System.Drawing.Size(120, 30)
        Me.btnMsgbox.Text = "MsgBox"
        Me.btnMsgbox.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btnMsgbox.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btnMsgbox.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btnMsgbox.Name = "btnMsgbox"
        Me.btnMsgbox.TabIndex = 0
        Me.Controls.Add(Me.btnMsgbox)
        Me.btnSave.Location = New System.Drawing.Point(250, 110)
        Me.btnSave.Size = New System.Drawing.Size(120, 30)
        Me.btnSave.Text = "Save"
        Me.btnSave.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btnSave.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btnSave.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 0
        Me.Controls.Add(Me.btnSave)
        Me.btnFont.Location = New System.Drawing.Point(250, 170)
        Me.btnFont.Size = New System.Drawing.Size(120, 30)
        Me.btnFont.Text = "Font"
        Me.btnFont.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btnFont.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btnFont.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btnFont.Name = "btnFont"
        Me.btnFont.TabIndex = 0
        Me.Controls.Add(Me.btnFont)
        Me.btnInput.Location = New System.Drawing.Point(250, 220)
        Me.btnInput.Size = New System.Drawing.Size(120, 30)
        Me.btnInput.Text = "Input Box"
        Me.btnInput.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btnInput.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btnInput.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btnInput.Name = "btnInput"
        Me.btnInput.TabIndex = 0
        Me.Controls.Add(Me.btnInput)
        Me.ClientSize = New System.Drawing.Size(640, 480)
        Me.Text = "Form1"
        Me.Name = "Form1"
        Me.ResumeLayout(False)
    End Sub
End Class
