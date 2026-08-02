' vybe-test: vb/vb_integer_overflow_checked_adv/test_vb_overflow_checked_convert_to_sbyte_large
' origin: languages/vb/tests/vb/test_vb_integer_overflow_checked_adv.rs

Module Program
    Sub Main()
        Try
            Dim i As Integer = 200
            Dim sb As SByte = Convert.ToSByte(i)
            Console.WriteLine(sb)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
