' vybe-test: vb/vb_system_exception_matrix/exception_invalid_cast_is_reported_as_type
' origin: languages/vb/tests/vb/test_vb_system_exception_matrix.rs

Imports System

Module M
    Sub Main()
        Try
            Dim value As Object = "text"
            Dim asInt As Integer = CInt(value)
            Console.WriteLine("ok")
        Catch ex As InvalidCastException
            Console.WriteLine(ex.GetType().Name)
        End Try
    End Sub
End Module
