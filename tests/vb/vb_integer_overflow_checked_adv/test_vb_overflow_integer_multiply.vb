' vybe-test: vb/vb_integer_overflow_checked_adv/test_vb_overflow_integer_multiply
' origin: languages/vb/tests/vb/test_vb_integer_overflow_checked_adv.rs

Module Program
    Sub Main()
        Try
            Dim x As Integer = 1000000
            Dim y As Integer = x * 1000000
            Console.WriteLine(y)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
