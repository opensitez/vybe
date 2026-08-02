' vybe-test: vb/vb_redim_preserve_valuetypes/redim_preserve_value_types
' origin: languages/vb/tests/vb/test_vb_redim_preserve_valuetypes.rs

Module M
    Sub Main()
        Dim arr() As Integer = {1, 2, 3}
        
        ' ReDim Preserve keeps existing values
        ReDim Preserve arr(4)
        
        arr(3) = 4
        arr(4) = 5
        
        For i As Integer = 0 To arr.Length - 1
            Console.WriteLine(arr(i))
        Next
    End Sub
End Module
