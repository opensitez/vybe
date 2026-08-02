' vybe-test: vb/vb_error_handling/try_catch_no_error
' origin: languages/vb/tests/vb/test_vb_error_handling.rs

Module M
    Sub Main()
        Try
            Console.WriteLine("ok")
        Catch ex As Exception
            Console.WriteLine("error")
        End Try
    End Sub
End Module
