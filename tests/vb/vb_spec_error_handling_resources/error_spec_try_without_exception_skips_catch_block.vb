' vybe-test: vb/vb_spec_error_handling_resources/error_spec_try_without_exception_skips_catch_block
' origin: languages/vb/tests/vb/test_vb_spec_error_handling_resources.rs

Module M
    Sub Main()
        Try
            Console.WriteLine("ok")
        Catch ex As Exception
            Console.WriteLine("catch")
        End Try
    End Sub
End Module
