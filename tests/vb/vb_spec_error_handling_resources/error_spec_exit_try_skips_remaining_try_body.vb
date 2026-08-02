' vybe-test: vb/vb_spec_error_handling_resources/error_spec_exit_try_skips_remaining_try_body
' origin: languages/vb/tests/vb/test_vb_spec_error_handling_resources.rs

Module M
    Sub Main()
        Try
            Console.WriteLine("before")
            Exit Try
            Console.WriteLine("after")
        Finally
            Console.WriteLine("finally")
        End Try
    End Sub
End Module
