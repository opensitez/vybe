' vybe-test: vb/vb_comprehensive/nested_for_loops
' origin: languages/vb/tests/vb/vb_comprehensive_test.rs

Module M
    Sub Main()
        Dim total As Integer = 0
        For i As Integer = 1 To 3
            For j As Integer = 1 To 3
                total = total + 1
            Next
        Next
        Console.WriteLine(total)
    End Sub
End Module
