' vybe-test: vb/vb_interop/b70_while_loop
' origin: languages/vb/tests/vb/vb_interop_test.rs

Dim n As Integer = 1
Dim result As Integer = 1
While n <= 5
    result = result * n
    n = n + 1
End While
Console.WriteLine(result)
