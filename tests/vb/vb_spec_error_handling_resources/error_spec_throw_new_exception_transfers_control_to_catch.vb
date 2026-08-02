' vybe-test: vb/vb_spec_error_handling_resources/error_spec_throw_new_exception_transfers_control_to_catch
' origin: languages/vb/tests/vb/test_vb_spec_error_handling_resources.rs

Module M
    Sub Main()
        Try
            Throw New Exception("x")
            Console.WriteLine("after")
        Catch ex As Exception
            Console.WriteLine("caught")
        End Try
    End Sub
End Module
