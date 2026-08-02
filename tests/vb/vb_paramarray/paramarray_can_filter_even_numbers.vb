' vybe-test: vb/vb_paramarray/paramarray_can_filter_even_numbers
' origin: languages/vb/tests/vb/test_vb_paramarray.rs

Module M
    Function CountEven(ParamArray values() As Integer) As Integer
        Dim total As Integer = 0
        For Each value As Integer In values
            If value Mod 2 = 0 Then
                total = total + 1
            End If
        Next
        Return total
    End Function

    Sub Main()
        Console.WriteLine(CountEven(1, 2, 4, 7, 8))
    End Sub
End Module
