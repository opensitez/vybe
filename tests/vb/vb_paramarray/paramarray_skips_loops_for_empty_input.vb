' vybe-test: vb/vb_paramarray/paramarray_skips_loops_for_empty_input
' origin: languages/vb/tests/vb/test_vb_paramarray.rs

Module M
    Function JoinAll(ParamArray values() As String) As String
        Dim result As String = "empty"
        For Each value As String In values
            result = result & value
        Next
        Return result
    End Function

    Sub Main()
        Console.WriteLine(JoinAll())
    End Sub
End Module
