' vybe-test: vb/vb_paramarray/paramarray_handles_negative_numbers
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
        Console.WriteLine(SumAll(10, -3, -2))
    End Sub
End Module
