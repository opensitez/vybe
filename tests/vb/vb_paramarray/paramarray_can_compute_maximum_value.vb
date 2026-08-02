' vybe-test: vb/vb_paramarray/paramarray_can_compute_maximum_value
' origin: languages/vb/tests/vb/test_vb_paramarray.rs

Module M
    Function MaxValue(ParamArray values() As Integer) As Integer
        Dim current As Integer = values(0)
        For Each value As Integer In values
            If value > current Then
                current = value
            End If
        Next
        Return current
    End Function

    Sub Main()
        Console.WriteLine(MaxValue(4, 9, 1, 7))
    End Sub
End Module
