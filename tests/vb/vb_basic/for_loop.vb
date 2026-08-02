' vybe-test: vb/vb_basic/for_loop
' origin: languages/vb/tests/vb/vb_basic_test.rs

Module Program
    Sub Main()
        Dim total As Integer = 0
        For i As Integer = 1 To 5
            total = total + i
        Next
        Console.WriteLine(total)
    End Sub
End Module
