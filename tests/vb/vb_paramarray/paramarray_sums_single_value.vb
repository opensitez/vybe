' vybe-test: vb/vb_paramarray/paramarray_sums_single_value
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
        Console.WriteLine(SumAll(7))
    End Sub
End Module
