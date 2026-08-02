' vybe-test: vb/vb_paramarray/paramarray_can_be_used_in_shared_methods
' origin: languages/vb/tests/vb/test_vb_paramarray.rs

Class MathBox
    Public Shared Function MultiplyAll(ParamArray values() As Integer) As Integer
        Dim total As Integer = 1
        For Each value As Integer In values
            total = total * value
        Next
        Return total
    End Function
End Class

Module M
    Sub Main()
        Console.WriteLine(MathBox.MultiplyAll(2, 3, 4))
    End Sub
End Module
