Public Class Form1
    Private Sub btnMsg_Click(sender As Object, e As EventArgs) Handles btnMsg.Click
        ' Msg Dialog
        MessageBox.Show("Hello", "Message")
    End Sub

    Private Sub btnFont_Click(sender As Object, e As EventArgs) Handles btnFont.Click
        ' Font Dialog
        Using fontDlg As New FontDialog()
            If fontDlg.ShowDialog() = DialogResult.OK Then
                ' In the original VB6, there was a txtBox referenced but not in the FRM
                ' I'll keep the logic but comment it out if txtBox doesn't exist or just show a msg
                ' txtBox.Font = fontDlg.Font
                MessageBox.Show("Font selected: " & fontDlg.Font.Name)
            End If
        End Using
    End Sub

    Private Sub btnColor_Click(sender As Object, e As EventArgs) Handles btnColor.Click
        ' Color Dialog
        Using colorDlg As New ColorDialog()
            If colorDlg.ShowDialog() = DialogResult.OK Then
                Me.BackColor = colorDlg.Color
            End If
        End Using
    End Sub

    Private Sub btnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click
        ' Open File Dialog
        Using openDlg As New OpenFileDialog()
            openDlg.Filter = "Text Files|*.txt|All Files|*.*"
            openDlg.Title = "Open File"
            If openDlg.ShowDialog() = DialogResult.OK Then
                Dim filename As String = openDlg.FileName
                MessageBox.Show("File selected: " & filename)
            End If
        End Using
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        ' Save File Dialog  
        Using saveDlg As New SaveFileDialog()
            saveDlg.Filter = "VB Files|*.vb|All Files|*.*"
            If saveDlg.ShowDialog() = DialogResult.OK Then
                Dim filename As String = saveDlg.FileName
                MessageBox.Show("Save to: " & filename)
            End If
        End Using
    End Sub

    Private Sub btnInput_Click(sender As Object, e As EventArgs) Handles btnInput.Click
        ' Input Box
        Dim name As String = InputBox("Enter your name:", "Name Entry", "Default")
        If Not String.IsNullOrEmpty(name) Then
            MessageBox.Show("Hello, " & name)
        End If
    End Sub

End Class
