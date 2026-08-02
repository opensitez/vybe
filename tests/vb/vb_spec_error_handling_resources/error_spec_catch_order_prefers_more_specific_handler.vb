' vybe-test: vb/vb_spec_error_handling_resources/error_spec_catch_order_prefers_more_specific_handler
' origin: languages/vb/tests/vb/test_vb_spec_error_handling_resources.rs

Module M
    Sub Main()
        Try
            Throw New ArgumentException("bad")
        Catch ex As ArgumentException
            Console.WriteLine("specific")
        Catch ex As Exception
            Console.WriteLine("general")
        End Try
    End Sub
End Module
