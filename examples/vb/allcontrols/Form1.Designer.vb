Partial Class Form1
    Inherits System.Windows.Forms.Form

    Friend WithEvents btn1 As System.Windows.Forms.Button
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents txt1 As System.Windows.Forms.TextBox
    Friend WithEvents chk1 As System.Windows.Forms.CheckBox
    Friend WithEvents opt1 As System.Windows.Forms.RadioButton
    Friend WithEvents cbo1 As System.Windows.Forms.ComboBox
    Friend WithEvents lst1 As System.Windows.Forms.ListBox
    Friend WithEvents fra1 As System.Windows.Forms.GroupBox
    Friend WithEvents pic1 As System.Windows.Forms.PictureBox
    Friend WithEvents rtf1 As System.Windows.Forms.RichTextBox
    Friend WithEvents web1 As System.Windows.Forms.WebBrowser
    Friend WithEvents tvw1 As System.Windows.Forms.TreeView
    Friend WithEvents pnl1 As System.Windows.Forms.Panel
    Friend WithEvents dgv1 As System.Windows.Forms.DataGridView
    Friend WithEvents lvw1 As System.Windows.Forms.ListView
    Friend WithEvents tab1 As System.Windows.Forms.TabControl
    Friend WithEvents pb1 As System.Windows.Forms.ProgressBar
    Friend WithEvents nud1 As System.Windows.Forms.NumericUpDown
    Friend WithEvents ms1 As System.Windows.Forms.MenuStrip
    Friend WithEvents cms1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ss1 As System.Windows.Forms.StatusStrip
    Friend WithEvents dtp1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents lnk1 As System.Windows.Forms.LinkLabel
    Friend WithEvents ts1 As System.Windows.Forms.ToolStrip
    Friend WithEvents trk1 As System.Windows.Forms.TrackBar
    Friend WithEvents mtxt1 As System.Windows.Forms.MaskedTextBox
    Friend WithEvents bs1 As System.Windows.Forms.BindingSource
    Friend WithEvents bnav1 As System.Windows.Forms.BindingNavigator
    Friend WithEvents hsb1 As System.Windows.Forms.HScrollBar
    Friend WithEvents vsb1 As System.Windows.Forms.VScrollBar
    Friend WithEvents ds1 As System.Data.DataSet
    Friend WithEvents dt1 As System.Data.DataTable
    Friend WithEvents da1 As System.Data.SqlClient.SqlDataAdapter
    Friend WithEvents mc1 As System.Windows.Forms.MonthCalendar
    Friend WithEvents sc1 As System.Windows.Forms.SplitContainer
    Friend WithEvents flp1 As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents tlp1 As System.Windows.Forms.TableLayoutPanel

    Private Sub InitializeComponent()
        Me.btn1 = New System.Windows.Forms.Button()
        Me.lbl1 = New System.Windows.Forms.Label()
        Me.txt1 = New System.Windows.Forms.TextBox()
        Me.chk1 = New System.Windows.Forms.CheckBox()
        Me.opt1 = New System.Windows.Forms.RadioButton()
        Me.cbo1 = New System.Windows.Forms.ComboBox()
        Me.lst1 = New System.Windows.Forms.ListBox()
        Me.fra1 = New System.Windows.Forms.GroupBox()
        Me.pic1 = New System.Windows.Forms.PictureBox()
        Me.rtf1 = New System.Windows.Forms.RichTextBox()
        Me.web1 = New System.Windows.Forms.WebBrowser()
        Me.tvw1 = New System.Windows.Forms.TreeView()
        Me.pnl1 = New System.Windows.Forms.Panel()
        Me.dgv1 = New System.Windows.Forms.DataGridView()
        Me.lvw1 = New System.Windows.Forms.ListView()
        Me.tab1 = New System.Windows.Forms.TabControl()
        Me.pb1 = New System.Windows.Forms.ProgressBar()
        Me.nud1 = New System.Windows.Forms.NumericUpDown()
        Me.ms1 = New System.Windows.Forms.MenuStrip()
        Me.cms1 = New System.Windows.Forms.ContextMenuStrip()
        Me.ss1 = New System.Windows.Forms.StatusStrip()
        Me.dtp1 = New System.Windows.Forms.DateTimePicker()
        Me.lnk1 = New System.Windows.Forms.LinkLabel()
        Me.ts1 = New System.Windows.Forms.ToolStrip()
        Me.trk1 = New System.Windows.Forms.TrackBar()
        Me.mtxt1 = New System.Windows.Forms.MaskedTextBox()
        Me.bs1 = New System.Windows.Forms.BindingSource()
        Me.bnav1 = New System.Windows.Forms.BindingNavigator()
        Me.hsb1 = New System.Windows.Forms.HScrollBar()
        Me.vsb1 = New System.Windows.Forms.VScrollBar()
        Me.ds1 = New System.Data.DataSet()
        Me.dt1 = New System.Data.DataTable()
        Me.da1 = New System.Data.SqlClient.SqlDataAdapter()
        Me.mc1 = New System.Windows.Forms.MonthCalendar()
        Me.sc1 = New System.Windows.Forms.SplitContainer()
        Me.flp1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.tlp1 = New System.Windows.Forms.TableLayoutPanel()
        Me.SuspendLayout()
        Me.btn1.Location = New System.Drawing.Point(20, 50)
        Me.btn1.Size = New System.Drawing.Size(120, 30)
        Me.btn1.Text = "btn1"
        Me.btn1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.btn1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.btn1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.btn1.Name = "btn1"
        Me.btn1.TabIndex = 0
        Me.Controls.Add(Me.btn1)
        Me.lbl1.Location = New System.Drawing.Point(30, 100)
        Me.lbl1.Size = New System.Drawing.Size(80, 20)
        Me.lbl1.Text = "lbl1"
        Me.lbl1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.lbl1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.lbl1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.TabIndex = 0
        Me.Controls.Add(Me.lbl1)
        Me.txt1.Location = New System.Drawing.Point(30, 130)
        Me.txt1.Size = New System.Drawing.Size(150, 25)
        Me.txt1.Text = ""
        Me.txt1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.txt1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.txt1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.txt1.Name = "txt1"
        Me.txt1.TabIndex = 0
        Me.Controls.Add(Me.txt1)
        Me.chk1.Location = New System.Drawing.Point(40, 190)
        Me.chk1.Size = New System.Drawing.Size(120, 20)
        Me.chk1.Text = "chk1"
        Me.chk1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.chk1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.chk1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.chk1.Name = "chk1"
        Me.chk1.TabIndex = 0
        Me.Controls.Add(Me.chk1)
        Me.opt1.Location = New System.Drawing.Point(50, 230)
        Me.opt1.Size = New System.Drawing.Size(120, 20)
        Me.opt1.Text = "opt1"
        Me.opt1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.opt1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.opt1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.opt1.Name = "opt1"
        Me.opt1.TabIndex = 0
        Me.Controls.Add(Me.opt1)
        Me.cbo1.Location = New System.Drawing.Point(20, 270)
        Me.cbo1.Size = New System.Drawing.Size(150, 25)
        Me.cbo1.Text = ""
        Me.cbo1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.cbo1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.cbo1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.cbo1.Name = "cbo1"
        Me.cbo1.TabIndex = 0
        Me.Controls.Add(Me.cbo1)
        Me.lst1.Location = New System.Drawing.Point(20, 310)
        Me.lst1.Size = New System.Drawing.Size(150, 100)
        Me.lst1.Text = ""
        Me.lst1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.lst1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.lst1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.lst1.Name = "lst1"
        Me.lst1.TabIndex = 0
        Me.Controls.Add(Me.lst1)
        Me.fra1.Location = New System.Drawing.Point(20, 430)
        Me.fra1.Size = New System.Drawing.Size(150, 90)
        Me.fra1.Text = "fra1"
        Me.fra1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.fra1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.fra1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.fra1.Name = "fra1"
        Me.fra1.TabIndex = 0
        Me.Controls.Add(Me.fra1)
        Me.pic1.Location = New System.Drawing.Point(20, 530)
        Me.pic1.Size = New System.Drawing.Size(150, 100)
        Me.pic1.Text = "pic1"
        Me.pic1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.pic1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.pic1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.pic1.Name = "pic1"
        Me.pic1.TabIndex = 0
        Me.Controls.Add(Me.pic1)
        Me.rtf1.Location = New System.Drawing.Point(20, 640)
        Me.rtf1.Size = New System.Drawing.Size(140, 70)
        Me.rtf1.Text = "Test"
        Me.rtf1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.rtf1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.rtf1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.rtf1.Name = "rtf1"
        Me.rtf1.TabIndex = 0
        Me.Controls.Add(Me.rtf1)
        Me.web1.Location = New System.Drawing.Point(200, 50)
        Me.web1.Size = New System.Drawing.Size(160, 110)
        Me.web1.Text = "web1"
        Me.web1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.web1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.web1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.web1.Name = "web1"
        Me.web1.TabIndex = 0
        Me.Controls.Add(Me.web1)
        Me.tvw1.Location = New System.Drawing.Point(200, 170)
        Me.tvw1.Size = New System.Drawing.Size(200, 100)
        Me.tvw1.Text = "tvw1"
        Me.tvw1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.tvw1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.tvw1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.tvw1.Name = "tvw1"
        Me.tvw1.TabIndex = 0
        Me.Controls.Add(Me.tvw1)
        Me.pnl1.Location = New System.Drawing.Point(200, 440)
        Me.pnl1.Size = New System.Drawing.Size(200, 150)
        Me.pnl1.Text = "pnl1"
        Me.pnl1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.pnl1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.pnl1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.pnl1.Name = "pnl1"
        Me.pnl1.TabIndex = 0
        Me.Controls.Add(Me.pnl1)
        Me.dgv1.Location = New System.Drawing.Point(200, 290)
        Me.dgv1.Size = New System.Drawing.Size(200, 140)
        Me.dgv1.Text = "dgv1"
        Me.dgv1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.dgv1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.dgv1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.dgv1.Name = "dgv1"
        Me.dgv1.TabIndex = 0
        Me.Controls.Add(Me.dgv1)
        Me.lvw1.Location = New System.Drawing.Point(410, 50)
        Me.lvw1.Size = New System.Drawing.Size(220, 150)
        Me.lvw1.Text = "lvw1"
        Me.lvw1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.lvw1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.lvw1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.lvw1.Name = "lvw1"
        Me.lvw1.TabIndex = 0
        Me.Controls.Add(Me.lvw1)
        Me.tab1.Location = New System.Drawing.Point(410, 210)
        Me.tab1.Size = New System.Drawing.Size(220, 200)
        Me.tab1.Text = "tab1"
        Me.tab1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.tab1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.tab1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.tab1.Name = "tab1"
        Me.tab1.TabIndex = 0
        Me.Controls.Add(Me.tab1)
        Me.pb1.Location = New System.Drawing.Point(410, 420)
        Me.pb1.Size = New System.Drawing.Size(200, 23)
        Me.pb1.Text = "pb1"
        Me.pb1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.pb1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.pb1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.pb1.Name = "pb1"
        Me.pb1.TabIndex = 0
        Me.Controls.Add(Me.pb1)
        Me.nud1.Location = New System.Drawing.Point(410, 450)
        Me.nud1.Size = New System.Drawing.Size(120, 23)
        Me.nud1.Text = "nud1"
        Me.nud1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.nud1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.nud1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.nud1.Name = "nud1"
        Me.nud1.TabIndex = 0
        Me.Controls.Add(Me.nud1)
        Me.ms1.Location = New System.Drawing.Point(410, 490)
        Me.ms1.Size = New System.Drawing.Size(220, 20)
        Me.ms1.Text = "ms1"
        Me.ms1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.ms1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.ms1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.ms1.Name = "ms1"
        Me.ms1.TabIndex = 0
        Me.Controls.Add(Me.ms1)
        Me.cms1.Location = New System.Drawing.Point(410, 520)
        Me.cms1.Size = New System.Drawing.Size(150, 24)
        Me.cms1.Text = "cms1"
        Me.cms1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.cms1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.cms1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.cms1.Name = "cms1"
        Me.cms1.TabIndex = 0
        Me.Controls.Add(Me.cms1)
        Me.ss1.Location = New System.Drawing.Point(410, 590)
        Me.ss1.Size = New System.Drawing.Size(210, 20)
        Me.ss1.Text = "ss1"
        Me.ss1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.ss1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.ss1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.ss1.Name = "ss1"
        Me.ss1.TabIndex = 0
        Me.Controls.Add(Me.ss1)
        Me.dtp1.Location = New System.Drawing.Point(410, 620)
        Me.dtp1.Size = New System.Drawing.Size(200, 23)
        Me.dtp1.Text = "dtp1"
        Me.dtp1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.dtp1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.dtp1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.dtp1.Name = "dtp1"
        Me.dtp1.TabIndex = 0
        Me.Controls.Add(Me.dtp1)
        Me.lnk1.Location = New System.Drawing.Point(530, 450)
        Me.lnk1.Size = New System.Drawing.Size(100, 20)
        Me.lnk1.Text = "lnk1"
        Me.lnk1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.lnk1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.lnk1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.lnk1.Name = "lnk1"
        Me.lnk1.TabIndex = 0
        Me.Controls.Add(Me.lnk1)
        Me.ts1.Location = New System.Drawing.Point(410, 660)
        Me.ts1.Size = New System.Drawing.Size(210, 20)
        Me.ts1.Text = "ts1"
        Me.ts1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.ts1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.ts1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.ts1.Name = "ts1"
        Me.ts1.TabIndex = 0
        Me.Controls.Add(Me.ts1)
        Me.trk1.Location = New System.Drawing.Point(410, 690)
        Me.trk1.Size = New System.Drawing.Size(200, 20)
        Me.trk1.Text = "trk1"
        Me.trk1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.trk1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.trk1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.trk1.Name = "trk1"
        Me.trk1.TabIndex = 0
        Me.Controls.Add(Me.trk1)
        Me.mtxt1.Location = New System.Drawing.Point(420, 720)
        Me.mtxt1.Size = New System.Drawing.Size(150, 23)
        Me.mtxt1.Text = ""
        Me.mtxt1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.mtxt1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.mtxt1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.mtxt1.Name = "mtxt1"
        Me.mtxt1.TabIndex = 0
        Me.Controls.Add(Me.mtxt1)
        Me.bs1.Name = "bs1"
        Me.bnav1.Location = New System.Drawing.Point(10, 850)
        Me.bnav1.Size = New System.Drawing.Size(200, 20)
        Me.bnav1.Text = "bnav1"
        Me.bnav1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.bnav1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.bnav1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.bnav1.Name = "bnav1"
        Me.bnav1.TabIndex = 0
        Me.Controls.Add(Me.bnav1)
        Me.hsb1.Location = New System.Drawing.Point(230, 850)
        Me.hsb1.Size = New System.Drawing.Size(170, 20)
        Me.hsb1.Text = "hsb1"
        Me.hsb1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.hsb1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.hsb1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.hsb1.Name = "hsb1"
        Me.hsb1.TabIndex = 0
        Me.Controls.Add(Me.hsb1)
        Me.vsb1.Location = New System.Drawing.Point(380, 50)
        Me.vsb1.Size = New System.Drawing.Size(10, 110)
        Me.vsb1.Text = "vsb1"
        Me.vsb1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.vsb1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.vsb1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.vsb1.Name = "vsb1"
        Me.vsb1.TabIndex = 0
        Me.Controls.Add(Me.vsb1)
        Me.ds1.DataSetName = "NewDataSet"
        Me.ds1.Name = "ds1"
        Me.dt1.TableName = "Table1"
        Me.dt1.Name = "dt1"
        Me.da1.Name = "da1"
        Me.mc1.Location = New System.Drawing.Point(410, 760)
        Me.mc1.Size = New System.Drawing.Size(220, 20)
        Me.mc1.Text = "mc1"
        Me.mc1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.mc1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.mc1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.mc1.Name = "mc1"
        Me.mc1.TabIndex = 0
        Me.Controls.Add(Me.mc1)
        Me.sc1.Location = New System.Drawing.Point(200, 600)
        Me.sc1.Size = New System.Drawing.Size(190, 80)
        Me.sc1.Text = "sc1"
        Me.sc1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.sc1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.sc1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.sc1.Name = "sc1"
        Me.sc1.TabIndex = 0
        Me.Controls.Add(Me.sc1)
        Me.flp1.Location = New System.Drawing.Point(200, 690)
        Me.flp1.Size = New System.Drawing.Size(200, 120)
        Me.flp1.Text = "flp1"
        Me.flp1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.flp1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.flp1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.flp1.Name = "flp1"
        Me.flp1.TabIndex = 0
        Me.Controls.Add(Me.flp1)
        Me.tlp1.Location = New System.Drawing.Point(20, 740)
        Me.tlp1.Size = New System.Drawing.Size(170, 100)
        Me.tlp1.Text = "tlp1"
        Me.tlp1.BackColor = System.Drawing.ColorTranslator.FromHtml("#f8fafc")
        Me.tlp1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#0f172a")
        Me.tlp1.Font = New System.Drawing.Font("Segoe UI", 12F)
        Me.tlp1.Name = "tlp1"
        Me.tlp1.TabIndex = 0
        Me.Controls.Add(Me.tlp1)
        Me.ClientSize = New System.Drawing.Size(640, 880)
        Me.Text = "Form1"
        Me.Name = "Form1"
        Me.ResumeLayout(False)
    End Sub
End Class
