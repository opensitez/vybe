' vybe-test: vb/vb_integer_overflow_checked_adv/test_vb_overflow_checked_short_multiplication
' origin: languages/vb/tests/vb/test_vb_integer_overflow_checked_adv.rs

Module Program
    Sub Main()
        Try
            Dim s1 As Short = 1000
            Dim s2 As Short = 100
            Dim s3 As Short = CShort(s1 * s2)
            Console.WriteLine(s3)
        Catch ex As OverflowException
            Console.WriteLine("OverflowException")
        End Try
    End Sub
End Module
