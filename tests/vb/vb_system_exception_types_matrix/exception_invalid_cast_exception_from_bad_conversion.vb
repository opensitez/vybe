' vybe-test: vb/vb_system_exception_types_matrix/exception_invalid_cast_exception_from_bad_conversion
' origin: languages/vb/tests/vb/test_vb_system_exception_types_matrix.rs

Module M
    Sub Main()
        Try
            Dim value As Object = "text"
            Dim typed As Integer = CInt(value)
            Console.WriteLine(typed)
        Catch ex As InvalidCastException
            Console.WriteLine(ex.GetType().Name)
        End Try
    End Sub
End Module
