' vybe-test: vb/vb_paramarray/paramarray_can_follow_required_prefix_argument
' origin: languages/vb/tests/vb/test_vb_paramarray.rs

Module M
    Function SumWithOffset(offset As Integer, ParamArray values() As Integer) As Integer
        Dim total As Integer = offset
        For Each value As Integer In values
            total = total + value
        Next
        Return total
    End Function

    Sub Main()
        Console.WriteLine(SumWithOffset(10, 1, 2, 3))
    End Sub
End Module
