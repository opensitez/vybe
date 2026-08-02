' vybe-test: vb/vb_system_boxing_unboxing_matrix/boxing_unboxing_directcast_throws_when_invalid
' origin: languages/vb/tests/vb/test_vb_system_boxing_unboxing_matrix.rs

Module M
    Sub Main()
        Dim boxed As Object = "hello"

        Try
            Dim x As Integer = DirectCast(boxed, Integer)
            Console.WriteLine("no")
        Catch ex As InvalidCastException
            Console.WriteLine("invalid")
        End Try
    End Sub
End Module
