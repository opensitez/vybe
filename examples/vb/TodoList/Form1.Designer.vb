Partial Class Form1
    Inherits System.Windows.Forms.Form

    Friend WithEvents lstTasks As System.Windows.Forms.ListBox
    Friend WithEvents txtTask As System.Windows.Forms.TextBox
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents lblCount As System.Windows.Forms.Label

    Private Sub InitializeComponent()
        Me.lstTasks = New System.Windows.Forms.ListBox()
        Me.txtTask = New System.Windows.Forms.TextBox()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.lblCount = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        Me.lstTasks.Location = New System.Drawing.Point(10, 100)
        Me.lstTasks.Size = New System.Drawing.Size(260, 140)
        Me.lstTasks.Text = ""
        Me.lstTasks.Name = "lstTasks"
        Me.lstTasks.TabIndex = 3
        Me.Controls.Add(Me.lstTasks)
        Me.txtTask.Location = New System.Drawing.Point(10, 60)
        Me.txtTask.Size = New System.Drawing.Size(180, 23)
        Me.txtTask.Text = ""
        Me.txtTask.Name = "txtTask"
        Me.txtTask.TabIndex = 1
        Me.Controls.Add(Me.txtTask)
        Me.btnAdd.Location = New System.Drawing.Point(200, 60)
        Me.btnAdd.Size = New System.Drawing.Size(75, 23)
        Me.btnAdd.Text = "Add"
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.TabIndex = 2
        Me.Controls.Add(Me.btnAdd)
        Me.btnRemove.Location = New System.Drawing.Point(10, 260)
        Me.btnRemove.Size = New System.Drawing.Size(85, 28)
        Me.btnRemove.Text = "Remove"
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.TabIndex = 4
        Me.Controls.Add(Me.btnRemove)
        Me.btnClear.Location = New System.Drawing.Point(100, 260)
        Me.btnClear.Size = New System.Drawing.Size(85, 28)
        Me.btnClear.Text = "Clear All"
        Me.btnClear.Name = "btnClear"
        Me.btnClear.TabIndex = 5
        Me.Controls.Add(Me.btnClear)
        Me.lblTitle.Location = New System.Drawing.Point(10, 40)
        Me.lblTitle.Size = New System.Drawing.Size(260, 20)
        Me.lblTitle.Text = "Todo List"
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.TabIndex = 0
        Me.Controls.Add(Me.lblTitle)
        Me.lblCount.Location = New System.Drawing.Point(190, 260)
        Me.lblCount.Size = New System.Drawing.Size(80, 20)
        Me.lblCount.Text = "Tasks: 0"
        Me.lblCount.Name = "lblCount"
        Me.lblCount.TabIndex = 6
        Me.Controls.Add(Me.lblCount)
        Me.ClientSize = New System.Drawing.Size(285, 300)
        Me.Text = "Todo List"
        Me.Name = "Form1"
        Me.ResumeLayout(False)
    End Sub
End Class
