' vybe-test: vb/vb_spec_error_handling_resources/error_spec_multiple_catch_blocks_can_select_specific_type
' origin: languages/vb/tests/vb/test_vb_spec_error_handling_resources.rs

Module M
    Sub Main()
        Try
            Throw New ArgumentException("bad")
        Catch ex As ArgumentException
            Console.WriteLine("arg")
        Catch ex As Exception
            Console.WriteLine("general")
        End Try
    End Sub
End Module
