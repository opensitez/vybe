<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
        Me.txtGrade = New System.Windows.Forms.TextBox()
        Me.btnGrade = New System.Windows.Forms.Button()
        Me.btnSum = New System.Windows.Forms.Button()
        Me.lbl1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'txtGrade
        '
        Me.txtGrade.Location = New System.Drawing.Point(50, 100)
        Me.txtGrade.Name = "txtGrade"
        Me.txtGrade.Size = New System.Drawing.Size(150, 20)
        Me.txtGrade.TabIndex = 0
        Me.txtGrade.Text = "85"
        '
        'btnGrade
        '
        Me.btnGrade.Location = New System.Drawing.Point(50, 150)
        Me.btnGrade.Name = "btnGrade"
        Me.btnGrade.Size = New System.Drawing.Size(140, 30)
        Me.btnGrade.TabIndex = 1
        Me.btnGrade.Text = "Calc Grade"
        Me.btnGrade.UseVisualStyleBackColor = True
        '
        'btnSum
        '
        Me.btnSum.Location = New System.Drawing.Point(50, 210)
        Me.btnSum.Name = "btnSum"
        Me.btnSum.Size = New System.Drawing.Size(150, 30)
        Me.btnSum.TabIndex = 2
        Me.btnSum.Text = "Sum Grades"
        Me.btnSum.UseVisualStyleBackColor = True
        '
        'lbl1
        '
        Me.lbl1.Location = New System.Drawing.Point(20, 60)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(280, 20)
        Me.lbl1.TabIndex = 3
        Me.lbl1.Text = "Shows use of select (Calc) and arrays (Sum)"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(340, 280)
        Me.Controls.Add(Me.lbl1)
        Me.Controls.Add(Me.btnSum)
        Me.Controls.Add(Me.btnGrade)
        Me.Controls.Add(Me.txtGrade)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtGrade As System.Windows.Forms.TextBox
    Friend WithEvents btnGrade As System.Windows.Forms.Button
    Friend WithEvents btnSum As System.Windows.Forms.Button
    Friend WithEvents lbl1 As System.Windows.Forms.Label
End Class
