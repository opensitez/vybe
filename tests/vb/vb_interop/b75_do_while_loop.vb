' vybe-test: vb/vb_interop/b75_do_while_loop
' origin: languages/vb/tests/vb/vb_interop_test.rs

Dim x As Integer = 0
Do While x < 3
    x = x + 1
Loop
Console.WriteLine(x)
