Partial Class Form1
    Inherits System.Windows.Forms.Form

    Friend WithEvents web1 As System.Windows.Forms.WebBrowser
    Friend WithEvents btn1 As System.Windows.Forms.Button

    Private Sub InitializeComponent()
        Me.web1 = New System.Windows.Forms.WebBrowser()
        Me.btn1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        Me.web1.Location = New System.Drawing.Point(10, 100)
        Me.web1.Size = New System.Drawing.Size(620, 570)
        Me.web1.Text = "Welcome This is an html web vddiew"
        Me.web1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.web1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.web1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.web1.Name = "web1"
        Me.web1.TabIndex = 0
        Me.web1.Enabled = True
        Me.web1.HTML = "<style>body { font-family: 'Inter', -apple-system, sans-serif; background: #0f172a; color: #f8fafc; display: flex; align-items: center; justify-content: center; height: 100vh; margin: 0; overflow: hidden; } .card { background: #1e293b; border: 1px solid #334155; padding: 48px; border-radius: 32px; box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5); text-align: center; max-width: 80%; } .icon { font-size: 64px; margin-bottom: 24px; animation: wave 2s infinite; } @keyframes wave { 0%, 100% { transform: rotate(0deg); } 50% { transform: rotate(15deg); } } h1 { font-size: 48px; font-weight: 800; margin: 0; color: #818cf8; letter-spacing: -1.5px; } .status { margin-top: 16px; padding: 8px 16px; background: rgba(129, 140, 248, 0.1); border-radius: 9999px; display: inline-block; color: #c7d2fe; font-weight: 600; font-size: 14px; } p { font-size: 18px; line-height: 1.6; color: #94a3b8; margin-top: 24px; }</style><div class=""card""><div class=""icon"">👋</div><h1>Vybe HTML</h1><div class=""status"">HTML Property: Synchronized</div><p>This is the Vybe HTML browser component.<br>The text you're reading is set via the <b>HTML property</b> in the designer!</p></div>"
        Me.web1.URL = "about:blank"
        Me.web1.Visible = True
        Me.btn1.Location = New System.Drawing.Point(70, 50)
        Me.btn1.Size = New System.Drawing.Size(120, 30)
        Me.btn1.Text = "Navigate"
        Me.btn1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btn1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btn1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btn1.Name = "btn1"
        Me.btn1.TabIndex = 0
        Me.btn1.Enabled = True
        Me.btn1.Visible = True
        Me.Controls.Add(Me.web1)
        Me.Controls.Add(Me.btn1)
        AddHandler Me.btn1.Click, AddressOf Me.btn1_Click
        AddHandler Me.Form1.Load, AddressOf Me.Form1_Load
        Me.ClientSize = New System.Drawing.Size(640, 680)
        Me.Text = "Form1"
        Me.Name = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub
End Class
