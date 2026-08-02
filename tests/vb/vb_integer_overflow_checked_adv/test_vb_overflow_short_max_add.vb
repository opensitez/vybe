' vybe-test: vb/vb_integer_overflow_checked_adv/test_vb_overflow_short_max_add
' origin: languages/vb/tests/vb/test_vb_integer_overflow_checked_adv.rs

Module Program
    Sub Main()
        Try
            Dim x As Short = Short.MaxValue
            Dim y As Short = CShort(x + 1)
            Console.WriteLine(y)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
