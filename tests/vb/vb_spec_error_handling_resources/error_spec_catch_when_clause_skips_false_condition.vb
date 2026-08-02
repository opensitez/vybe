' vybe-test: vb/vb_spec_error_handling_resources/error_spec_catch_when_clause_skips_false_condition
' origin: languages/vb/tests/vb/test_vb_spec_error_handling_resources.rs

Module M
    Sub Main()
        Try
            Throw New Exception("boom")
        Catch ex As Exception When ex.Message = "other"
            Console.WriteLine("matched")
        Catch ex As Exception
            Console.WriteLine("fallback")
        End Try
    End Sub
End Module
