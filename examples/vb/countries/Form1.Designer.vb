Partial Class Form1
    Inherits System.Windows.Forms.Form

    Friend WithEvents da1 As System.Data.SqlClient.SqlDataAdapter
    Friend WithEvents bs1 As System.Windows.Forms.BindingSource
    Friend WithEvents bnav1 As System.Windows.Forms.BindingNavigator
    Friend WithEvents txt1 As System.Windows.Forms.TextBox
    Friend WithEvents txt2 As System.Windows.Forms.TextBox
    Friend WithEvents txt3 As System.Windows.Forms.TextBox
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents lbl2 As System.Windows.Forms.Label
    Friend WithEvents lbl3 As System.Windows.Forms.Label
    Friend WithEvents lbl4 As System.Windows.Forms.Label
    Friend WithEvents dgv1 As System.Windows.Forms.DataGridView

    Private Sub InitializeComponent()
        Me.da1 = New System.Data.SqlClient.SqlDataAdapter()
        Me.bs1 = New System.Windows.Forms.BindingSource()
        Me.bnav1 = New System.Windows.Forms.BindingNavigator()
        Me.txt1 = New System.Windows.Forms.TextBox()
        Me.txt2 = New System.Windows.Forms.TextBox()
        Me.txt3 = New System.Windows.Forms.TextBox()
        Me.lbl1 = New System.Windows.Forms.Label()
        Me.lbl2 = New System.Windows.Forms.Label()
        Me.lbl3 = New System.Windows.Forms.Label()
        Me.lbl4 = New System.Windows.Forms.Label()
        Me.dgv1 = New System.Windows.Forms.DataGridView()
        Me.SuspendLayout()
        Me.da1.ConnectionString = "Data Source=friends.db"
        Me.da1.Name = "da1"
        Me.da1.BackColor = "#f8fafc"
        Me.da1.DbName = "friends"
        Me.da1.DbPassword = "vybe"
        Me.da1.DbPort = "3306"
        Me.da1.DbType = "MySQL"
        Me.da1.DbUser = "vybe"
        Me.da1.Font = "Segoe UI, 12px"
        Me.da1.ForeColor = "#0f172a"
        Me.bs1.DataSource = Me.da1
        Me.bs1.DataMember = "Countries"
        Me.bs1.Name = "bs1"
        Me.bs1.BackColor = "#f8fafc"
        Me.bs1.Font = "Segoe UI, 12px"
        Me.bs1.ForeColor = "#0f172a"
        Me.bnav1.Location = New System.Drawing.Point(30, 440)
        Me.bnav1.Size = New System.Drawing.Size(300, 25)
        Me.bnav1.Text = "bnav1"
        Me.bnav1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.bnav1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.bnav1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.bnav1.Name = "bnav1"
        Me.bnav1.BindingSource = Me.bs1
        Me.bnav1.TabIndex = 0
        Me.bnav1.Enabled = True
        Me.bnav1.Visible = True
        Me.txt1.Location = New System.Drawing.Point(140, 110)
        Me.txt1.Size = New System.Drawing.Size(150, 25)
        Me.txt1.Text = ""
        Me.txt1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.txt1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.txt1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.txt1.Name = "txt1"
        Me.txt1.DataBindings.Add("Text", Me.bs1, "Language")
        Me.txt1.TabIndex = 0
        Me.txt1.Enabled = True
        Me.txt1.Visible = True
        Me.txt2.Location = New System.Drawing.Point(140, 160)
        Me.txt2.Size = New System.Drawing.Size(150, 25)
        Me.txt2.Text = ""
        Me.txt2.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.txt2.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.txt2.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.txt2.Name = "txt2"
        Me.txt2.DataBindings.Add("Text", Me.bs1, "ISOCOde")
        Me.txt2.TabIndex = 0
        Me.txt2.Enabled = True
        Me.txt2.Visible = True
        Me.txt3.Location = New System.Drawing.Point(140, 210)
        Me.txt3.Size = New System.Drawing.Size(150, 25)
        Me.txt3.Text = ""
        Me.txt3.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.txt3.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.txt3.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.txt3.Name = "txt3"
        Me.txt3.DataBindings.Add("Text", Me.bs1, "CountryName")
        Me.txt3.TabIndex = 0
        Me.txt3.Enabled = True
        Me.txt3.Visible = True
        Me.lbl1.Location = New System.Drawing.Point(20, 110)
        Me.lbl1.Size = New System.Drawing.Size(80, 20)
        Me.lbl1.Text = "Language"
        Me.lbl1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.lbl1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.lbl1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.TabIndex = 0
        Me.lbl1.Visible = True
        Me.lbl2.Location = New System.Drawing.Point(20, 170)
        Me.lbl2.Size = New System.Drawing.Size(80, 20)
        Me.lbl2.Text = "ISO Code"
        Me.lbl2.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.lbl2.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.lbl2.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.lbl2.Name = "lbl2"
        Me.lbl2.TabIndex = 0
        Me.lbl2.Visible = True
        Me.lbl3.Location = New System.Drawing.Point(20, 210)
        Me.lbl3.Size = New System.Drawing.Size(100, 20)
        Me.lbl3.Text = "Country Name"
        Me.lbl3.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.lbl3.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.lbl3.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.lbl3.Name = "lbl3"
        Me.lbl3.TabIndex = 0
        Me.lbl3.Visible = True
        Me.lbl4.Location = New System.Drawing.Point(30, 60)
        Me.lbl4.Size = New System.Drawing.Size(540, 20)
        Me.lbl4.Text = "Simple Demo of data bound controls in Vybe interpreter"
        Me.lbl4.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.lbl4.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.lbl4.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.lbl4.Name = "lbl4"
        Me.lbl4.TabIndex = 0
        Me.lbl4.Visible = True
        Me.dgv1.Location = New System.Drawing.Point(320, 110)
        Me.dgv1.Size = New System.Drawing.Size(360, 300)
        Me.dgv1.Text = "dgv1"
        Me.dgv1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.dgv1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.dgv1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.dgv1.Name = "dgv1"
        Me.dgv1.DataSource = Me.bs1
        Me.dgv1.DataMember = "countryid"
        Me.dgv1.TabIndex = 0
        Me.dgv1.AllowUserToAddRows = True
        Me.dgv1.AllowUserToDeleteRows = True
        Me.dgv1.Enabled = True
        Me.dgv1.Visible = True
        Me.Controls.Add(Me.bnav1)
        Me.Controls.Add(Me.txt1)
        Me.Controls.Add(Me.txt2)
        Me.Controls.Add(Me.txt3)
        Me.Controls.Add(Me.lbl1)
        Me.Controls.Add(Me.lbl2)
        Me.Controls.Add(Me.lbl3)
        Me.Controls.Add(Me.lbl4)
        Me.Controls.Add(Me.dgv1)
        Me.ClientSize = New System.Drawing.Size(740, 480)
        Me.Text = "Form1"
        Me.Name = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub
End Class
