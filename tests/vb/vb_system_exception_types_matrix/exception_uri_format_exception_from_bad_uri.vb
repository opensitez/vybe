' vybe-test: vb/vb_system_exception_types_matrix/exception_uri_format_exception_from_bad_uri
' origin: languages/vb/tests/vb/test_vb_system_exception_types_matrix.rs

Imports System

Module M
    Sub Main()
        Try
            Dim url As New Uri("://definitely-not-valid")
            Console.WriteLine(url.AbsoluteUri)
        Catch ex As UriFormatException
            Console.WriteLine(ex.GetType().Name)
        End Try
    End Sub
End Module
