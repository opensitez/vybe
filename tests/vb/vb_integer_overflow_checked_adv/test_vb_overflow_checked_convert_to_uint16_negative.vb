' vybe-test: vb/vb_integer_overflow_checked_adv/test_vb_overflow_checked_convert_to_uint16_negative
' origin: languages/vb/tests/vb/test_vb_integer_overflow_checked_adv.rs

Module Program
    Sub Main()
        Try
            Dim i As Integer = -50
            Dim us As UShort = Convert.ToUInt16(i)
            Console.WriteLine(us)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
