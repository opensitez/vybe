' vybe-test: vb/vb_integer_overflow_checked_adv/test_vb_overflow_checked_convert_to_byte_negative
' origin: languages/vb/tests/vb/test_vb_integer_overflow_checked_adv.rs

Module Program
    Sub Main()
        Try
            Dim i As Integer = -1
            Dim b As Byte = Convert.ToByte(i)
            Console.WriteLine(b)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
