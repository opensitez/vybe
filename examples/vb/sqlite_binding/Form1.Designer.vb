Partial Class Form1
    Inherits System.Windows.Forms.Form

    Friend WithEvents dgvContacts As System.Windows.Forms.DataGridView
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents lblEmail As System.Windows.Forms.Label
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents lblPhone As System.Windows.Forms.Label
    Friend WithEvents txtPhone As System.Windows.Forms.TextBox
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents bs1 As System.Windows.Forms.BindingSource
    Friend WithEvents da1 As System.Data.SqlClient.SqlDataAdapter

    Private Sub InitializeComponent()
        Me.dgvContacts = New System.Windows.Forms.DataGridView()
        Me.lblName = New System.Windows.Forms.Label()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.lblEmail = New System.Windows.Forms.Label()
        Me.txtEmail = New System.Windows.Forms.TextBox()
        Me.lblPhone = New System.Windows.Forms.Label()
        Me.txtPhone = New System.Windows.Forms.TextBox()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.bs1 = New System.Windows.Forms.BindingSource()
        Me.da1 = New System.Data.SqlClient.SqlDataAdapter()
        Me.SuspendLayout()
        Me.dgvContacts.Location = New System.Drawing.Point(20, 45)
        Me.dgvContacts.Size = New System.Drawing.Size(450, 200)
        Me.dgvContacts.Text = "dgvContacts"
        Me.dgvContacts.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.dgvContacts.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.dgvContacts.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.dgvContacts.Name = "dgvContacts"
        Me.dgvContacts.DataSource = Me.bs1
        Me.dgvContacts.TabIndex = 1
        Me.Controls.Add(Me.dgvContacts)
        Me.lblName.Location = New System.Drawing.Point(20, 260)
        Me.lblName.Size = New System.Drawing.Size(60, 20)
        Me.lblName.Text = "Name:"
        Me.lblName.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.lblName.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.lblName.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.lblName.Name = "lblName"
        Me.lblName.TabIndex = 2
        Me.Controls.Add(Me.lblName)
        Me.txtName.Location = New System.Drawing.Point(85, 258)
        Me.txtName.Size = New System.Drawing.Size(180, 25)
        Me.txtName.Text = ""
        Me.txtName.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.txtName.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.txtName.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.txtName.Name = "txtName"
        Me.txtName.DataBindings.Add("Text", Me.bs1, "Name")
        Me.txtName.TabIndex = 3
        Me.Controls.Add(Me.txtName)
        Me.lblEmail.Location = New System.Drawing.Point(20, 290)
        Me.lblEmail.Size = New System.Drawing.Size(60, 20)
        Me.lblEmail.Text = "Email:"
        Me.lblEmail.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.lblEmail.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.lblEmail.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.TabIndex = 4
        Me.Controls.Add(Me.lblEmail)
        Me.txtEmail.Location = New System.Drawing.Point(85, 288)
        Me.txtEmail.Size = New System.Drawing.Size(180, 25)
        Me.txtEmail.Text = ""
        Me.txtEmail.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.txtEmail.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.txtEmail.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.DataBindings.Add("Text", Me.bs1, "Email")
        Me.txtEmail.TabIndex = 5
        Me.Controls.Add(Me.txtEmail)
        Me.lblPhone.Location = New System.Drawing.Point(280, 260)
        Me.lblPhone.Size = New System.Drawing.Size(50, 20)
        Me.lblPhone.Text = "Phone:"
        Me.lblPhone.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.lblPhone.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.lblPhone.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.lblPhone.Name = "lblPhone"
        Me.lblPhone.TabIndex = 6
        Me.Controls.Add(Me.lblPhone)
        Me.txtPhone.Location = New System.Drawing.Point(335, 258)
        Me.txtPhone.Size = New System.Drawing.Size(135, 25)
        Me.txtPhone.Text = ""
        Me.txtPhone.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.txtPhone.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.txtPhone.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.txtPhone.Name = "txtPhone"
        Me.txtPhone.DataBindings.Add("Text", Me.bs1, "Phone")
        Me.txtPhone.TabIndex = 7
        Me.Controls.Add(Me.txtPhone)
        Me.btnAdd.Location = New System.Drawing.Point(20, 325)
        Me.btnAdd.Size = New System.Drawing.Size(100, 30)
        Me.btnAdd.Text = "Add Contact"
        Me.btnAdd.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btnAdd.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btnAdd.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.TabIndex = 8
        Me.Controls.Add(Me.btnAdd)
        Me.btnDelete.Location = New System.Drawing.Point(130, 325)
        Me.btnDelete.Size = New System.Drawing.Size(110, 30)
        Me.btnDelete.Text = "Delete Selected"
        Me.btnDelete.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btnDelete.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btnDelete.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.TabIndex = 9
        Me.Controls.Add(Me.btnDelete)
        Me.btnRefresh.Location = New System.Drawing.Point(250, 325)
        Me.btnRefresh.Size = New System.Drawing.Size(80, 30)
        Me.btnRefresh.Text = "Refresh"
        Me.btnRefresh.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btnRefresh.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btnRefresh.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.TabIndex = 10
        Me.Controls.Add(Me.btnRefresh)
        Me.lblStatus.Location = New System.Drawing.Point(20, 365)
        Me.lblStatus.Size = New System.Drawing.Size(450, 20)
        Me.lblStatus.Text = "Ready"
        Me.lblStatus.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.lblStatus.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.lblStatus.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.TabIndex = 11
        Me.Controls.Add(Me.lblStatus)
        Me.lblTitle.Location = New System.Drawing.Point(10, 420)
        Me.lblTitle.Size = New System.Drawing.Size(460, 20)
        Me.lblTitle.Text = "SQLite Contacts - Data Binding Demo"
        Me.lblTitle.BackColor = System.Drawing.ColorTranslator.FromHtml("#000000")
        Me.lblTitle.ForeColor = System.Drawing.ColorTranslator.FromHtml("#ffffff")
        Me.lblTitle.Font = New System.Drawing.Font("Segoe UI", 14F)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.TabIndex = 0
        Me.Controls.Add(Me.lblTitle)
        Me.bs1.DataSource = Me.da1
        Me.bs1.DataMember = "contacts"
        Me.bs1.Name = "bs1"
        Me.da1.SelectCommand = "SELECT * FROM contacts ORDER BY Name"
        Me.da1.ConnectionString = "Data Source=contacts.db"
        Me.da1.Name = "da1"
        Me.ClientSize = New System.Drawing.Size(490, 455)
        Me.Text = "SQLite Contacts"
        Me.Name = "Form1"
        Me.BackColor = System.Drawing.ColorTranslator.FromHtml("#d6d6d6")
        Me.ResumeLayout(False)
    End Sub
End Class
