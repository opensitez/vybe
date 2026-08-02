' vybe-test: vb/vb_integer_overflow_checked_adv/test_vb_overflow_ulong_max_add
' origin: languages/vb/tests/vb/test_vb_integer_overflow_checked_adv.rs

Module Program
    Sub Main()
        Try
            Dim ul As ULong = ULong.MaxValue
            Dim ul2 As ULong = ul + 1UL
            Console.WriteLine(ul2)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
