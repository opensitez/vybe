Partial Class SocketForm
    Inherits System.Windows.Forms.Form

    Friend WithEvents Size As System.Drawing.Size
    Friend WithEvents lblServer As System.Windows.Forms.Label
    Friend WithEvents btnListen As System.Windows.Forms.Button
    Friend WithEvents txtServerLog As System.Windows.Forms.TextBox
    Friend WithEvents lblClient As System.Windows.Forms.Label
    Friend WithEvents btnConnect As System.Windows.Forms.Button
    Friend WithEvents txtMessage As System.Windows.Forms.TextBox
    Friend WithEvents btnSend As System.Windows.Forms.Button
    Friend WithEvents txtClientLog As System.Windows.Forms.TextBox

    Private Sub InitializeComponent()
        Me.Size = New System.Drawing.Size()
        Me.lblServer = New System.Windows.Forms.Label()
        Me.btnListen = New System.Windows.Forms.Button()
        Me.txtServerLog = New System.Windows.Forms.TextBox()
        Me.lblClient = New System.Windows.Forms.Label()
        Me.btnConnect = New System.Windows.Forms.Button()
        Me.txtMessage = New System.Windows.Forms.TextBox()
        Me.btnSend = New System.Windows.Forms.Button()
        Me.txtClientLog = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        Me.Size.Location = New System.Drawing.Point(290, 40)
        Me.Size.Size = New System.Drawing.Size(100, 100)
        Me.Size.Text = "Size"
        Me.Size.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.Size.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.Size.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.Size.Name = "Size"
        Me.Size.TabIndex = 0
        Me.Size.Enabled = True
        Me.Size.Visible = True
        Me.lblServer.Location = New System.Drawing.Point(20, 70)
        Me.lblServer.Size = New System.Drawing.Size(180, 20)
        Me.lblServer.Text = "TCP Server (Port 8080)"
        Me.lblServer.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.lblServer.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.lblServer.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.lblServer.Name = "lblServer"
        Me.lblServer.TabIndex = 0
        Me.lblServer.Visible = True
        Me.btnListen.Location = New System.Drawing.Point(20, 100)
        Me.btnListen.Size = New System.Drawing.Size(120, 30)
        Me.btnListen.Text = "Start Listening"
        Me.btnListen.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btnListen.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btnListen.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btnListen.Name = "btnListen"
        Me.btnListen.TabIndex = 0
        Me.btnListen.Enabled = True
        Me.btnListen.Visible = True
        Me.txtServerLog.Location = New System.Drawing.Point(20, 140)
        Me.txtServerLog.Size = New System.Drawing.Size(350, 350)
        Me.txtServerLog.Text = ""
        Me.txtServerLog.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.txtServerLog.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.txtServerLog.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.txtServerLog.Name = "txtServerLog"
        Me.txtServerLog.TabIndex = 0
        Me.txtServerLog.Enabled = True
        Me.txtServerLog.Multiline = True
        Me.txtServerLog.ReadOnly = True
        Me.txtServerLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtServerLog.Visible = True
        Me.lblClient.Location = New System.Drawing.Point(410, 70)
        Me.lblClient.Size = New System.Drawing.Size(200, 20)
        Me.lblClient.Text = "TCP Client"
        Me.lblClient.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.lblClient.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.lblClient.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.lblClient.Name = "lblClient"
        Me.lblClient.TabIndex = 0
        Me.lblClient.Visible = True
        Me.btnConnect.Location = New System.Drawing.Point(410, 100)
        Me.btnConnect.Size = New System.Drawing.Size(120, 30)
        Me.btnConnect.Text = "Connect to Server"
        Me.btnConnect.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btnConnect.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btnConnect.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btnConnect.Name = "btnConnect"
        Me.btnConnect.TabIndex = 0
        Me.btnConnect.Enabled = True
        Me.btnConnect.Visible = True
        Me.txtMessage.Location = New System.Drawing.Point(410, 140)
        Me.txtMessage.Size = New System.Drawing.Size(320, 30)
        Me.txtMessage.Text = "Hello Server!"
        Me.txtMessage.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.txtMessage.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.txtMessage.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.txtMessage.Name = "txtMessage"
        Me.txtMessage.TabIndex = 0
        Me.txtMessage.Enabled = True
        Me.txtMessage.Visible = True
        Me.btnSend.Location = New System.Drawing.Point(640, 100)
        Me.btnSend.Size = New System.Drawing.Size(80, 25)
        Me.btnSend.Text = "Send"
        Me.btnSend.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btnSend.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btnSend.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btnSend.Name = "btnSend"
        Me.btnSend.TabIndex = 0
        Me.btnSend.Enabled = False
        Me.btnSend.Visible = True
        Me.txtClientLog.Location = New System.Drawing.Point(410, 180)
        Me.txtClientLog.Size = New System.Drawing.Size(320, 310)
        Me.txtClientLog.Text = ""
        Me.txtClientLog.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.txtClientLog.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.txtClientLog.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.txtClientLog.Name = "txtClientLog"
        Me.txtClientLog.TabIndex = 0
        Me.txtClientLog.Enabled = True
        Me.txtClientLog.Multiline = True
        Me.txtClientLog.ReadOnly = True
        Me.txtClientLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtClientLog.Visible = True
        Me.Controls.Add(Me.Size)
        Me.Controls.Add(Me.lblServer)
        Me.Controls.Add(Me.btnListen)
        Me.Controls.Add(Me.txtServerLog)
        Me.Controls.Add(Me.lblClient)
        Me.Controls.Add(Me.btnConnect)
        Me.Controls.Add(Me.txtMessage)
        Me.Controls.Add(Me.btnSend)
        Me.Controls.Add(Me.txtClientLog)
        AddHandler Me.btnListen.Click, AddressOf Me.btnListen_Click
        AddHandler Me.btnConnect.Click, AddressOf Me.btnConnect_Click
        AddHandler Me.btnSend.Click, AddressOf Me.btnSend_Click
        Me.ClientSize = New System.Drawing.Size(780, 580)
        Me.Text = "Socket Communication Demo"
        Me.Name = "SocketForm"
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub
End Class
