' vybe-test: vb/vb_system_exception_types_matrix/exception_overflow_exception_from_narrowing_conversion
' origin: languages/vb/tests/vb/test_vb_system_exception_types_matrix.rs

Module M
    Sub Main()
        Try
            Dim value As Byte = CByte(300)
            Console.WriteLine(value)
        Catch ex As OverflowException
            Console.WriteLine(ex.GetType().Name)
        End Try
    End Sub
End Module
