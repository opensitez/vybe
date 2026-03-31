Imports System.Collections.Generic
Imports System.Drawing
Imports System.Windows.Forms

Public Class Connect4Form
    Private Const ROWS As Integer = 6
    Private Const COLS As Integer = 7
    ' 0 = Empty, 1 = Player (Red), 2 = AI (Yellow)
    Private _board(ROWS - 1, COLS - 1) As Integer
    Private _panels(ROWS - 1, COLS - 1) As Panel
    Private _isPlayerTurn As Boolean = True
    Private _isGameOver As Boolean = False

    Private Sub Connect4Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Console.WriteLine("DEBUG: Connect4Form_Load called")
        InitializeBoard()
    End Sub

    Private Sub InitializeBoard()
        Console.WriteLine("DEBUG: InitializeBoard start")
        pnlBoard.Controls.Clear()
        pnlBoard.RowCount = 6
        pnlBoard.ColumnCount = 7

        For r As Integer = 0 To 5
            For c As Integer = 0 To 6
                Dim p As New Panel()
                p.BackColor = Color.White
                p.BorderStyle = BorderStyle.FixedSingle
                p.Margin = New Padding(2)
                p.Dock = DockStyle.Fill
                p.Tag = r & "," & c
                
                Console.WriteLine("DEBUG: Adding panel at " & r & "," & c)
                pnlBoard.Controls.Add(p, c, r)
                _panels(r, c) = p
            Next
        Next
        Console.WriteLine("DEBUG: InitializeBoard done")
    End Sub

    Private Sub Cell_Click(sender As Object, e As EventArgs)
        If Not _isPlayerTurn Or _isGameOver Then Return
        
        Dim p = DirectCast(sender, Panel)
        Dim coords = DirectCast(p.Tag, Point)
        Dim col = coords.Y
        
        If MakeMove(col, 1) Then
            If _isGameOver Then Return
            _isPlayerTurn = False
            lblStatus.Text = "AI Thinking..."
            Application.DoEvents()
            
            ' Small delay for realism
            System.Threading.Thread.Sleep(500)
            
            Dim aiCol = GetBestMove()
            MakeMove(aiCol, 2)
            _isPlayerTurn = True
            If Not _isGameOver Then
                lblStatus.Text = "Your Turn (Red)"
            End If
        End If
    End Sub

    Private Function MakeMove(col As Integer, player As Integer) As Boolean
        ' Find lowest available row
        For r As Integer = ROWS - 1 To 0 Step -1
            If _board(r, col) = 0 Then
                _board(r, col) = player
                _panels(r, col).BackColor = If(player = 1, Color.Red, Color.Yellow)
                
                If CheckWin(r, col, player) Then
                    _isGameOver = True
                    lblStatus.Text = If(player = 1, "YOU WIN!", "AI WINS!")
                ElseIf IsBoardFull() Then
                    _isGameOver = True
                    lblStatus.Text = "DRAW!"
                End If
                Return True
            End If
        Next
        Return False
    End Function

    Private Function CheckWin(row As Integer, col As Integer, player As Integer) As Boolean
        ' Horizontal
        Dim count As Integer = 0
        For c As Integer = 0 To COLS - 1
            If _board(row, c) = player Then
                count += 1
                If count >= 4 Then Return True
            Else
                count = 0
            End If
        Next
        
        ' Vertical
        count = 0
        For r As Integer = 0 To ROWS - 1
            If _board(r, col) = player Then
                count += 1
                If count >= 4 Then Return True
            Else
                count = 0
            End If
        Next
        
        ' Diagonal 1 (\)
        count = 0
        Dim r1 = row, c1 = col
        While r1 > 0 And c1 > 0
            r1 -= 1 : c1 -= 1
        End While
        While r1 < ROWS And c1 < COLS
            If _board(r1, c1) = player Then
                count += 1
                If count >= 4 Then Return True
            Else
                count = 0
            End If
            r1 += 1 : c1 += 1
        End While
        
        ' Diagonal 2 (/)
        count = 0
        Dim r2 = row, c2 = col
        While r2 < ROWS - 1 And c2 > 0
            r2 += 1 : c2 -= 1
        End While
        While r2 >= 0 And c2 < COLS
            If _board(r2, c2) = player Then
                count += 1
                If count >= 4 Then Return True
            Else
                count = 0
            End If
            r2 -= 1 : c2 += 1
        End While
        
        Return False
    End Function

    Private Function IsBoardFull() As Boolean
        For c As Integer = 0 To COLS - 1
            If _board(0, c) = 0 Then Return False
        Next
        Return True
    End Function

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        _isGameOver = False
        _isPlayerTurn = True
        lblStatus.Text = "Your Turn (Red)"
        InitializeBoard()
    End Sub

    ' AI Logic: Minimax with Alpha-Beta
    Private Function GetBestMove() As Integer
        Dim bestScore As Integer = -1000000
        Dim bestCol As Integer = 0
        
        ' Simple lookahead
        For c As Integer = 0 To COLS - 1
            If _board(0, c) = 0 Then
                Dim r = DropTemp(c, 2)
                Dim score = Minimax(4, -1000000, 1000000, False)
                UndoTemp(r, c)
                
                If score > bestScore Then
                    bestScore = score
                    bestCol = c
                End If
            End If
        Next
        Return bestCol
    End Function

    Private Function Minimax(depth As Integer, alpha As Integer, beta As Integer, isMaximizing As Boolean) As Integer
        If depth = 0 Or _isGameOver Then
            Return EvaluateBoard()
        End If
        
        If isMaximizing Then
            Dim maxEval As Integer = -1000000
            For c As Integer = 0 To COLS - 1
                If _board(0, c) = 0 Then
                    Dim r = DropTemp(c, 2)
                    Dim eval = Minimax(depth - 1, alpha, beta, False)
                    UndoTemp(r, c)
                    maxEval = Math.Max(maxEval, eval)
                    alpha = Math.Max(alpha, eval)
                    If beta <= alpha Then Exit For
                End If
            Next
            Return maxEval
        Else
            Dim minEval As Integer = 1000000
            For c As Integer = 0 To COLS - 1
                If _board(0, c) = 0 Then
                    Dim r = DropTemp(c, 1)
                    Dim eval = Minimax(depth - 1, alpha, beta, True)
                    UndoTemp(r, c)
                    minEval = Math.Min(minEval, eval)
                    beta = Math.Min(beta, eval)
                    if beta <= alpha Then Exit For
                End If
            Next
            Return minEval
        End If
    End Function

    Private Function DropTemp(col As Integer, player As Integer) As Integer
        For r As Integer = ROWS - 1 To 0 Step -1
            If _board(r, col) = 0 Then
                _board(r, col) = player
                Return r
            End If
        Next
        Return -1
    End Function

    Private Sub UndoTemp(row As Integer, col As Integer)
        _board(row, col) = 0
    End Sub

    ' Very basic heuristic
    Private Function EvaluateBoard() As Integer
        Dim score As Integer = 0
        ' Favor center column
        For r As Integer = 0 To ROWS - 1
            If _board(r, 3) = 2 Then score += 3
            If _board(r, 3) = 1 Then score -= 3
        Next
        ' Note: Proper Connect4 AI would check for potential 3-in-a-rows etc.
        ' This is a simplified version.
        Return score
    End Function

End Class
