Public Class Form1
    Dim dblLastNum As Double    ' Stores the first number
    Dim strOp As String         ' Stores the operator (+, -, /, *)
    Dim blnClearDisplay As Boolean ' Flag to clear text when a new number is typed

    Private Sub AddDigit(strDigit As String)
        If blnClearDisplay Then
            txtCalc.Text = strDigit
            blnClearDisplay = False
        Else
            txtCalc.Text = txtCalc.Text & strDigit
        End If
    End Sub

    Private Sub btn0_Click(sender As Object, e As EventArgs) Handles btn0.Click
        AddDigit("0")
    End Sub

    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
        AddDigit("1")
    End Sub

    Private Sub btn2_Click(sender As Object, e As EventArgs) Handles btn2.Click
        AddDigit("2")
    End Sub

    Private Sub btn3_Click(sender As Object, e As EventArgs) Handles btn3.Click
        AddDigit("3")
    End Sub

    Private Sub btn4_Click(sender As Object, e As EventArgs) Handles btn4.Click
        AddDigit("4")
    End Sub

    Private Sub btn5_Click(sender As Object, e As EventArgs) Handles btn5.Click
        AddDigit("5")
    End Sub

    Private Sub btn6_Click(sender As Object, e As EventArgs) Handles btn6.Click
        AddDigit("6")
    End Sub

    Private Sub btn7_Click(sender As Object, e As EventArgs) Handles btn7.Click
        AddDigit("7")
    End Sub

    Private Sub btn8_Click(sender As Object, e As EventArgs) Handles btn8.Click
        AddDigit("8")
    End Sub

    Private Sub btn9_Click(sender As Object, e As EventArgs) Handles btn9.Click
        AddDigit("9")
    End Sub

    Private Sub SetOperator(Op As String)
        dblLastNum = Val(txtCalc.Text)
        strOp = Op
        blnClearDisplay = True
    End Sub

    Private Sub btnPlus_Click(sender As Object, e As EventArgs) Handles btnPlus.Click
        SetOperator("+")
    End Sub

    Private Sub btnMinus_Click(sender As Object, e As EventArgs) Handles btnMinus.Click
        SetOperator("-")
    End Sub

    Private Sub btnTimes_Click(sender As Object, e As EventArgs) Handles btnTimes.Click
        SetOperator("*")
    End Sub

    Private Sub btnDiv_Click(sender As Object, e As EventArgs) Handles btnDiv.Click
        SetOperator("/")
    End Sub

    Private Sub btnEquals_Click(sender As Object, e As EventArgs) Handles btnEquals.Click
        Dim dblCurrentNum As Double
        dblCurrentNum = Val(txtCalc.Text)

        Select Case strOp
            Case "+"
                txtCalc.Text = (dblLastNum + dblCurrentNum).ToString()
            Case "-"
                txtCalc.Text = (dblLastNum - dblCurrentNum).ToString()
            Case "*"
                txtCalc.Text = (dblLastNum * dblCurrentNum).ToString()
            Case "/"
                If dblCurrentNum <> 0 Then
                    txtCalc.Text = (dblLastNum / dblCurrentNum).ToString()
                Else
                    MessageBox.Show("Error: Division by zero", "Calculator", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
        End Select

        ' Reset state so the next number typed starts fresh
        blnClearDisplay = True
    End Sub

    Private Sub btnDot_Click(sender As Object, e As EventArgs) Handles btnDot.Click
        ' Only add a dot if there isn't one already
        If Not txtCalc.Text.Contains(".") Then
            txtCalc.Text = txtCalc.Text & "."
        End If
    End Sub

    Private Sub btnClearDisplay_Click(sender As Object, e As EventArgs) Handles btnClearDisplay.Click
        txtCalc.Clear()
        dblLastNum = 0
        strOp = ""
        blnClearDisplay = False
    End Sub

End Class
