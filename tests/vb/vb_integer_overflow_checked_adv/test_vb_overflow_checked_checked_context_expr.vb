' vybe-test: vb/vb_integer_overflow_checked_adv/test_vb_overflow_checked_checked_context_expr
' origin: languages/vb/tests/vb/test_vb_integer_overflow_checked_adv.rs

Module Program
    Sub Main()
        Try
            Dim a As Integer = Integer.MaxValue - 5
            Dim b As Integer = 10
            Dim c As Integer = a + b
            Console.WriteLine(c)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
