' vybe-test: vb/vb_integer_overflow_checked_adv/test_vb_overflow_long_max_add
' origin: languages/vb/tests/vb/test_vb_integer_overflow_checked_adv.rs

Module Program
    Sub Main()
        Try
            Dim l As Long = Long.MaxValue
            Dim l2 As Long = l + 1L
            Console.WriteLine(l2)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
