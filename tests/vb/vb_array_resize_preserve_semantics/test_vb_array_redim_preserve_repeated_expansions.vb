' vybe-test: vb/vb_array_resize_preserve_semantics/test_vb_array_redim_preserve_repeated_expansions
' origin: languages/vb/tests/vb/test_vb_array_resize_preserve_semantics.rs

Module Program
    Sub Main()
        Dim arr(0) As Integer
        arr(0) = 1
        For i As Integer = 1 To 4
            ReDim Preserve arr(i)
            arr(i) = (i + 1) * 10
        Next
        Console.WriteLine(String.Join(",", arr))
    End Sub
End Module
