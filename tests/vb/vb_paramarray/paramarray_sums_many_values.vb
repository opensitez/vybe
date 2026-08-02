' vybe-test: vb/vb_paramarray/paramarray_sums_many_values
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
        Console.WriteLine(SumAll(1, 2, 3, 4, 5))
    End Sub
End Module
