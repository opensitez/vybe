' vybe-test: vb/vb_integer_overflow_checked_adv/test_vb_overflow_narrowing_double_to_integer
' origin: languages/vb/tests/vb/test_vb_integer_overflow_checked_adv.rs

Module Program
    Sub Main()
        Try
            Dim d As Double = 1e15
            Dim i As Integer = CInt(d)
            Console.WriteLine(i)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
