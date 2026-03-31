Public Class Form1
    Dim turn As String
    Dim gameOver As Boolean

    Private Sub Form1_Load() Handles Me.Load
        ResetGame()
    End Sub

    Private Sub ResetGame()
        turn = "X"
        gameOver = False
        lblStatus.Text = "Your Turn (X)"
        
        btn0.Text = ""
        btn1.Text = ""
        btn2.Text = ""
        btn3.Text = ""
        btn4.Text = ""
        btn5.Text = ""
        btn6.Text = ""
        btn7.Text = ""
        btn8.Text = ""
        
        btn0.Enabled = True
        btn1.Enabled = True
        btn2.Enabled = True
        btn3.Enabled = True
        btn4.Enabled = True
        btn5.Enabled = True
        btn6.Enabled = True
        btn7.Enabled = True
        btn8.Enabled = True
    End Sub

    Private Sub btn0_Click() Handles btn0.Click
        HandleClick(btn0)
    End Sub

    Private Sub btn1_Click() Handles btn1.Click
        HandleClick(btn1)
    End Sub

    Private Sub btn2_Click() Handles btn2.Click
        HandleClick(btn2)
    End Sub

    Private Sub btn3_Click() Handles btn3.Click
        HandleClick(btn3)
    End Sub

    Private Sub btn4_Click() Handles btn4.Click
        HandleClick(btn4)
    End Sub

    Private Sub btn5_Click() Handles btn5.Click
        HandleClick(btn5)
    End Sub

    Private Sub btn6_Click() Handles btn6.Click
        HandleClick(btn6)
    End Sub

    Private Sub btn7_Click() Handles btn7.Click
        HandleClick(btn7)
    End Sub

    Private Sub btn8_Click() Handles btn8.Click
        HandleClick(btn8)
    End Sub

    Private Sub HandleClick(btn As Object)
        If gameOver Then
            Exit Sub
        End If
        If turn <> "X" Then
            Exit Sub
        End If
        If btn.Text <> "" Then
            Exit Sub
        End If
        
        btn.Text = "X"
        CheckWin()
        
        If Not gameOver Then
            turn = "O"
            lblStatus.Text = "Computer thinking..."
            ComputerMove()
        End If
    End Sub

    Private Sub ComputerMove()
        If gameOver Then
            Exit Sub
        End If

        ' Try to win
        If TryMove("O") Then
            CheckWin()
            If Not gameOver Then
                turn = "X"
                lblStatus.Text = "Your Turn (X)"
            End If
            Exit Sub
        End If

        ' Try to block player
        If TryMove("X") Then
            CheckWin()
            If Not gameOver Then
                turn = "X"
                lblStatus.Text = "Your Turn (X)"
            End If
            Exit Sub
        End If

        ' Take center if free
        If btn4.Text = "" Then
            btn4.Text = "O"
            CheckWin()
            If Not gameOver Then
                turn = "X"
                lblStatus.Text = "Your Turn (X)"
            End If
            Exit Sub
        End If

        ' Take a corner
        If btn0.Text = "" Then
            btn0.Text = "O"
            CheckWin()
            If Not gameOver Then
                turn = "X"
                lblStatus.Text = "Your Turn (X)"
            End If
            Exit Sub
        End If
        If btn2.Text = "" Then
            btn2.Text = "O"
            CheckWin()
            If Not gameOver Then
                turn = "X"
                lblStatus.Text = "Your Turn (X)"
            End If
            Exit Sub
        End If
        If btn6.Text = "" Then
            btn6.Text = "O"
            CheckWin()
            If Not gameOver Then
                turn = "X"
                lblStatus.Text = "Your Turn (X)"
            End If
            Exit Sub
        End If
        If btn8.Text = "" Then
            btn8.Text = "O"
            CheckWin()
            If Not gameOver Then
                turn = "X"
                lblStatus.Text = "Your Turn (X)"
            End If
            Exit Sub
        End If

        ' Take any open spot
        If btn1.Text = "" Then
            btn1.Text = "O"
        ElseIf btn3.Text = "" Then
            btn3.Text = "O"
        ElseIf btn5.Text = "" Then
            btn5.Text = "O"
        ElseIf btn7.Text = "" Then
            btn7.Text = "O"
        End If

        CheckWin()
        If Not gameOver Then
            turn = "X"
            lblStatus.Text = "Your Turn (X)"
        End If
    End Sub

    ' Try to find a winning move for the given mark, and place "O" there
    Private Function TryMove(mark As String) As Boolean
        ' Check all 8 lines for two-in-a-row with one empty
        If TryLine(btn0, btn1, btn2, mark) Then Return True
        If TryLine(btn3, btn4, btn5, mark) Then Return True
        If TryLine(btn6, btn7, btn8, mark) Then Return True
        If TryLine(btn0, btn3, btn6, mark) Then Return True
        If TryLine(btn1, btn4, btn7, mark) Then Return True
        If TryLine(btn2, btn5, btn8, mark) Then Return True
        If TryLine(btn0, btn4, btn8, mark) Then Return True
        If TryLine(btn2, btn4, btn6, mark) Then Return True
        Return False
    End Function

    ' If two cells have 'mark' and one is empty, place "O" in the empty one
    Private Function TryLine(b1 As Object, b2 As Object, b3 As Object, mark As String) As Boolean
        If b1.Text = mark And b2.Text = mark And b3.Text = "" Then
            b3.Text = "O"
            Return True
        End If
        If b1.Text = mark And b3.Text = mark And b2.Text = "" Then
            b2.Text = "O"
            Return True
        End If
        If b2.Text = mark And b3.Text = mark And b1.Text = "" Then
            b1.Text = "O"
            Return True
        End If
        Return False
    End Function

    Private Sub CheckWin()
        If CheckLine(btn0, btn1, btn2) Then
            Exit Sub
        End If
        If CheckLine(btn3, btn4, btn5) Then
            Exit Sub
        End If
        If CheckLine(btn6, btn7, btn8) Then
            Exit Sub
        End If
        
        If CheckLine(btn0, btn3, btn6) Then
            Exit Sub
        End If
        If CheckLine(btn1, btn4, btn7) Then
            Exit Sub
        End If
        If CheckLine(btn2, btn5, btn8) Then
            Exit Sub
        End If
        
        If CheckLine(btn0, btn4, btn8) Then
            Exit Sub
        End If
        If CheckLine(btn2, btn4, btn6) Then
            Exit Sub
        End If
        
        If btn0.Text <> "" And btn1.Text <> "" And btn2.Text <> "" And btn3.Text <> "" And btn4.Text <> "" And btn5.Text <> "" And btn6.Text <> "" And btn7.Text <> "" And btn8.Text <> "" Then
             lblStatus.Text = "It's a Draw!"
             gameOver = True
        End If
    End Sub

    Private Function CheckLine(b1 As Object, b2 As Object, b3 As Object) As Boolean
        If b1.Text <> "" And b1.Text = b2.Text And b2.Text = b3.Text Then
            If b1.Text = "X" Then
                lblStatus.Text = "You Win!"
            Else
                lblStatus.Text = "Computer Wins!"
            End If
            gameOver = True
            DisableAll()
            Return True
        End If
        Return False
    End Function

    Private Sub DisableAll()
        btn0.Enabled = False
        btn1.Enabled = False
        btn2.Enabled = False
        btn3.Enabled = False
        btn4.Enabled = False
        btn5.Enabled = False
        btn6.Enabled = False
        btn7.Enabled = False
        btn8.Enabled = False
    End Sub

    Private Sub btnReset_Click() Handles btnReset.Click
        ResetGame()
    End Sub
End Class
