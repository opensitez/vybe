Module Program
    Sub Main()
        Dim board(8) As Char
        ResetBoard(board)
        Dim current As Char = "X"c

        Do
            PlayGame(board, current)
            Console.Write("Play again? (y/n): ")
            Dim ans = Console.ReadLine()
            If String.IsNullOrEmpty(ans) OrElse ans.ToLower() <> "y" Then Exit Do
            ResetBoard(board)
            current = "X"c
        Loop
    End Sub

    Sub ResetBoard(ByRef board() As Char)
        For i As Integer = 0 To 8
            board(i) = " "c
        Next
    End Sub

    Sub PlayGame(ByRef board() As Char, ByRef current As Char)
        Dim moves As Integer = 0
        While True
            DrawBoard(board)
            Console.Write($"Player {current}, enter move (1-9): ")
            Dim input = Console.ReadLine()
            Dim pos As Integer
            If Integer.TryParse(input, pos) AndAlso pos >= 1 AndAlso pos <= 9 AndAlso board(pos - 1) = " "c Then
                board(pos - 1) = current
                moves += 1
                If CheckWin(board, current) Then
                    DrawBoard(board)
                    Console.WriteLine($"Player {current} wins!")
                    Exit While
                End If
                If moves >= 9 Then
                    DrawBoard(board)
                    Console.WriteLine("It's a draw.")
                    Exit While
                End If
                current = If(current = "X"c, "O"c, "X"c)
            Else
                Console.WriteLine("Invalid move, try again.")
            End If
        End While
    End Sub

    Sub DrawBoard(board() As Char)
        Console.Clear()
        Console.WriteLine($" {board(0)} | {board(1)} | {board(2)} ")
        Console.WriteLine("---+---+---")
        Console.WriteLine($" {board(3)} | {board(4)} | {board(5)} ")
        Console.WriteLine("---+---+---")
        Console.WriteLine($" {board(6)} | {board(7)} | {board(8)} ")
        Console.WriteLine()
    End Sub

    Function CheckWin(b() As Char, p As Char) As Boolean
        Dim wins As Integer()() = New Integer()() {
            New Integer() {0, 1, 2},
            New Integer() {3, 4, 5},
            New Integer() {6, 7, 8},
            New Integer() {0, 3, 6},
            New Integer() {1, 4, 7},
            New Integer() {2, 5, 8},
            New Integer() {0, 4, 8},
            New Integer() {2, 4, 6}
        }
        For Each w In wins
            If b(w(0)) = p AndAlso b(w(1)) = p AndAlso b(w(2)) = p Then Return True
        Next
        Return False
    End Function
End Module
