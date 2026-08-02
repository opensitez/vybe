' vybe-test: vb/vb_spec_error_handling_resources/error_spec_nested_try_inner_catch_handles_exception
' origin: languages/vb/tests/vb/test_vb_spec_error_handling_resources.rs

Module M
    Sub Main()
        Try
            Try
                Throw New Exception("inner")
            Catch ex As Exception
                Console.WriteLine("caught inner")
            End Try
        Catch ex As Exception
            Console.WriteLine("outer")
        End Try
    End Sub
End Module
