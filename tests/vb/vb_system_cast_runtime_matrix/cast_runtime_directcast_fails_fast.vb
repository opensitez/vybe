' vybe-test: vb/vb_system_cast_runtime_matrix/cast_runtime_directcast_fails_fast
' origin: languages/vb/tests/vb/test_vb_system_cast_runtime_matrix.rs

Class A
End Class

Class B
Inherits A
End Class

Module M
    Sub Main()
        Dim b As A = New A()

        Try
            Dim c As B = DirectCast(b, B)
            Console.WriteLine("bad")
        Catch ex As Exception
            Console.WriteLine(TypeOf ex Is InvalidCastException)
        End Try
    End Sub
End Module
