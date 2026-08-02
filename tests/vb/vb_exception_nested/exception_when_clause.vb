' vybe-test: vb/vb_exception_nested/exception_when_clause
' origin: languages/vb/tests/vb/test_vb_exception_nested.rs

Module M
    Sub Main()
        Dim code = 404
        Try
            Throw New Exception("Error")
        Catch ex As Exception When code = 200
            Console.WriteLine("OK")
        Catch ex As Exception When code = 404
            Console.WriteLine("Not Found")
        End Try
    End Sub
End Module
