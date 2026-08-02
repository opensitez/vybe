' vybe-test: vb/vb_integer_overflow_checked_adv/test_vb_overflow_checked_decimal_max_multiply
' origin: languages/vb/tests/vb/test_vb_integer_overflow_checked_adv.rs

Module Program
    Sub Main()
        Try
            Dim d As Decimal = Decimal.MaxValue
            Dim d2 As Decimal = d * 2D
            Console.WriteLine(d2)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
