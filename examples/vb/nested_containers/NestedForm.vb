Public Class NestedForm
    Inherits System.Windows.Forms.Form

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        ' Root Panel
        Me.pnlRoot = New System.Windows.Forms.Panel()
        Me.lblRoot = New System.Windows.Forms.Label()
        Me.pnlNested1 = New System.Windows.Forms.Panel()
        Me.lblNested1 = New System.Windows.Forms.Label()
        Me.btnNested1 = New System.Windows.Forms.Button()
        Me.pnlNested2 = New System.Windows.Forms.Panel()
        Me.lblNested2 = New System.Windows.Forms.Label()
        Me.chkNested2 = New System.Windows.Forms.CheckBox()
        
        Me.SuspendLayout()
        
        ' pnlRoot
        Me.pnlRoot.Name = "pnlRoot"
        Me.pnlRoot.Location = New System.Drawing.Point(50, 50)
        Me.pnlRoot.Size = New System.Drawing.Size(700, 500)
        Me.pnlRoot.BackColor = System.Drawing.Color.FromArgb(200, 200, 200) ' Gray
        Me.pnlRoot.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlRoot.Controls.Add(Me.lblRoot)
        Me.pnlRoot.Controls.Add(Me.pnlNested1)

        ' lblRoot
        Me.lblRoot.Name = "lblRoot"
        Me.lblRoot.Text = "Root Panel (Gray)"
        Me.lblRoot.Location = New System.Drawing.Point(10, 10)
        Me.lblRoot.Size = New System.Drawing.Size(200, 20)

        ' pnlNested1
        Me.pnlNested1.Name = "pnlNested1"
        Me.pnlNested1.Location = New System.Drawing.Point(50, 50)
        Me.pnlNested1.Size = New System.Drawing.Size(300, 400)
        Me.pnlNested1.BackColor = System.Drawing.Color.FromArgb(173, 216, 230) ' LightBlue
        Me.pnlNested1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlNested1.Controls.Add(Me.lblNested1)
        Me.pnlNested1.Controls.Add(Me.btnNested1)
        Me.pnlNested1.Controls.Add(Me.pnlNested2)

        ' lblNested1
        Me.lblNested1.Name = "lblNested1"
        Me.lblNested1.Text = "Nested Panel 1 (Blue)"
        Me.lblNested1.Location = New System.Drawing.Point(10, 10)
        Me.lblNested1.Size = New System.Drawing.Size(200, 20)

        ' btnNested1
        Me.btnNested1.Name = "btnNested1"
        Me.btnNested1.Text = "Click Me"
        Me.btnNested1.Location = New System.Drawing.Point(10, 50)
        Me.btnNested1.Size = New System.Drawing.Size(100, 30)
        
        ' pnlNested2
        Me.pnlNested2.Name = "pnlNested2"
        Me.pnlNested2.Location = New System.Drawing.Point(50, 150)
        Me.pnlNested2.Size = New System.Drawing.Size(200, 200)
        Me.pnlNested2.BackColor = System.Drawing.Color.FromArgb(255, 182, 193) ' LightPink
        Me.pnlNested2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlNested2.Controls.Add(Me.lblNested2)
        Me.pnlNested2.Controls.Add(Me.chkNested2)

        ' lblNested2
        Me.lblNested2.Name = "lblNested2"
        Me.lblNested2.Text = "Deep Nested (Pink)"
        Me.lblNested2.Location = New System.Drawing.Point(10, 10)
        Me.lblNested2.Size = New System.Drawing.Size(150, 20)

        ' chkNested2
        Me.chkNested2.Name = "chkNested2"
        Me.chkNested2.Text = "Deep CheckBox"
        Me.chkNested2.Location = New System.Drawing.Point(10, 50)
        Me.chkNested2.Size = New System.Drawing.Size(150, 20)
        
        ' NestedForm
        Me.Name = "NestedForm"
        Me.Text = "Nested Container Test"
        Me.ClientSize = New System.Drawing.Size(800, 600)
        Me.Controls.Add(Me.pnlRoot)
        Me.ResumeLayout(False)
    End Sub

    Friend WithEvents pnlRoot As System.Windows.Forms.Panel
    Friend WithEvents lblRoot As System.Windows.Forms.Label
    Friend WithEvents pnlNested1 As System.Windows.Forms.Panel
    Friend WithEvents lblNested1 As System.Windows.Forms.Label
    Friend WithEvents btnNested1 As System.Windows.Forms.Button
    Friend WithEvents pnlNested2 As System.Windows.Forms.Panel
    Friend WithEvents lblNested2 As System.Windows.Forms.Label
    Friend WithEvents chkNested2 As System.Windows.Forms.CheckBox

    Private Sub btnNested1_Click(sender As Object, e As EventArgs) Handles btnNested1.Click
        MsgBox("Button inside Blue Panel clicked!")
    End Sub
End Class
