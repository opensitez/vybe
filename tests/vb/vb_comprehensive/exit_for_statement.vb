' vybe-test: vb/vb_comprehensive/exit_for_statement
' origin: languages/vb/tests/vb/vb_comprehensive_test.rs

Module M
    Sub Main()
        For i As Integer = 1 To 10
            If i = 4 Then Exit For
            Console.WriteLine(i)
        Next
    End Sub
End Module
