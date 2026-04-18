Public Class Form1
    Private Sub btnGrade_Click(sender As Object, e As EventArgs) Handles btnGrade.Click
        ' Test Select Case
        Dim grade As Integer
        If Integer.TryParse(txtGrade.Text, grade) Then
            Select Case grade
                Case 90 To 100
                    MessageBox.Show("Grade: A - Excellent!", "Grade Result")
                Case 80 To 89
                    MessageBox.Show("Grade: B - Good work!", "Grade Result")
                Case 70 To 79
                    MessageBox.Show("Grade: C", "Grade Result")
                Case 60 To 69
                    MessageBox.Show("Grade: D", "Grade Result")
                Case Else
                    MessageBox.Show("Grade: F", "Grade Result")
            End Select
        Else
            MessageBox.Show("Please enter a valid numeric grade.", "Error")
        End If
    End Sub

    Private Sub btnSum_Click(sender As Object, e As EventArgs) Handles btnSum.Click
        Dim numbers() As Integer = {10, 20, 30, 40, 50}
        
        ' Test array with loop
        Dim total As Integer = 0

        For i As Integer = 0 To numbers.Length - 1
            total += numbers(i)
        Next

        MessageBox.Show("Sum of all numbers: " & total.ToString(), "Sum Result")
    End Sub
End Class
