' vybe-test: vb/vb_comprehensive/for_each_with_accumulation
' origin: languages/vb/tests/vb/vb_comprehensive_test.rs

Module M
    Sub Main()
        Dim nums() As Integer = {1, 2, 3, 4, 5}
        Dim total As Integer = 0
        For Each n As Integer In nums
            total = total + n
        Next
        Console.WriteLine(total)
    End Sub
End Module
