' vybe-test: vb/vb_paramarray/paramarray_handles_zero_values_without_special_cases
' origin: languages/vb/tests/vb/test_vb_paramarray.rs

Module M
    Function SumAll(ParamArray values() As Integer) As Integer
        Dim total As Integer = 0
        For Each value As Integer In values
            total = total + value
        Next
        Return total
    End Function

    Sub Main()
        Console.WriteLine(SumAll(0, 0, 0))
    End Sub
End Module
