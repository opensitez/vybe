' vybe-test: vb/vb_integer_overflow_checked_adv/test_vb_overflow_byte_max_add
' origin: languages/vb/tests/vb/test_vb_integer_overflow_checked_adv.rs

Module Program
    Sub Main()
        Try
            Dim b As Byte = Byte.MaxValue
            Dim b2 As Byte = CByte(b + 1)
            Console.WriteLine(b2)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
