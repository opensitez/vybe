' vybe-test: vb/vb_interop/b68_for_loop
' origin: languages/vb/tests/vb/vb_interop_test.rs

Dim total As Integer = 0
For i As Integer = 1 To 5
    total = total + i
Next
Console.WriteLine(total)
