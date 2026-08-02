' vybe-test: vb/vb_integer_overflow_checked_adv/test_vb_overflow_checked_math_abs_min
' origin: languages/vb/tests/vb/test_vb_integer_overflow_checked_adv.rs

Module Program
    Sub Main()
        Try
            Dim x As Integer = Integer.MinValue
            Dim y As Integer = Math.Abs(x)
            Console.WriteLine(y)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
