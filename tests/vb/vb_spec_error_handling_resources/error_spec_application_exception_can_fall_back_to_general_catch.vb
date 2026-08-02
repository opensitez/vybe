' vybe-test: vb/vb_spec_error_handling_resources/error_spec_application_exception_can_fall_back_to_general_catch
' origin: languages/vb/tests/vb/test_vb_spec_error_handling_resources.rs

Module M
    Sub Main()
        Try
            Throw New ApplicationException("boom")
        Catch ex As ArgumentException
            Console.WriteLine("arg")
        Catch ex As Exception
            Console.WriteLine("general")
        End Try
    End Sub
End Module
