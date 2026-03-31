Partial Class PaintForm
    Inherits System.Windows.Forms.Form

    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.pnlHeader = New System.Windows.Forms.Panel()
        Me.btnDoodle = New System.Windows.Forms.Button()
        Me.btnCircle = New System.Windows.Forms.Button()
        Me.btnRectangle = New System.Windows.Forms.Button()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.pnlCanvas = New System.Windows.Forms.Panel()
        Me.pnlHeader.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlHeader
        '
        Me.pnlHeader.Controls.Add(Me.btnClear)
        Me.pnlHeader.Controls.Add(Me.btnRectangle)
        Me.pnlHeader.Controls.Add(Me.btnCircle)
        Me.pnlHeader.Controls.Add(Me.btnDoodle)
        Me.pnlHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlHeader.Name = "pnlHeader"
        Me.pnlHeader.Size = New System.Drawing.Size(800, 50)
        Me.pnlHeader.TabIndex = 0
        '
        'btnDoodle
        '
        Me.btnDoodle.Location = New System.Drawing.Point(12, 12)
        Me.btnDoodle.Name = "btnDoodle"
        Me.btnDoodle.Size = New System.Drawing.Size(75, 23)
        Me.btnDoodle.TabIndex = 0
        Me.btnDoodle.Text = "Doodle"
        Me.btnDoodle.UseVisualStyleBackColor = True
        '
        'btnCircle
        '
        Me.btnCircle.Location = New System.Drawing.Point(93, 12)
        Me.btnCircle.Name = "btnCircle"
        Me.btnCircle.Size = New System.Drawing.Size(75, 23)
        Me.btnCircle.TabIndex = 1
        Me.btnCircle.Text = "Circle"
        Me.btnCircle.UseVisualStyleBackColor = True
        '
        'btnRectangle
        '
        Me.btnRectangle.Location = New System.Drawing.Point(174, 12)
        Me.btnRectangle.Name = "btnRectangle"
        Me.btnRectangle.Size = New System.Drawing.Size(75, 23)
        Me.btnRectangle.TabIndex = 2
        Me.btnRectangle.Text = "Rectangle"
        Me.btnRectangle.UseVisualStyleBackColor = True
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(713, 12)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(75, 23)
        Me.btnClear.TabIndex = 3
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'pnlCanvas
        '
        Me.pnlCanvas.BackColor = System.Drawing.Color.White
        Me.pnlCanvas.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlCanvas.Location = New System.Drawing.Point(0, 50)
        Me.pnlCanvas.Name = "pnlCanvas"
        Me.pnlCanvas.Size = New System.Drawing.Size(800, 400)
        Me.pnlCanvas.TabIndex = 1
        '
        'PaintForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.pnlCanvas)
        Me.Controls.Add(Me.pnlHeader)
        Me.Name = "PaintForm"
        Me.Text = "VB.NET Paint Sample"
        Me.pnlHeader.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pnlHeader As Panel
    Friend WithEvents btnDoodle As Button
    Friend WithEvents btnCircle As Button
    Friend WithEvents btnRectangle As Button
    Friend WithEvents btnClear As Button
    Friend WithEvents pnlCanvas As Panel
End Class
