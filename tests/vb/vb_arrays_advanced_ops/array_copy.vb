' vybe-test: vb/vb_arrays_advanced_ops/array_copy
' origin: languages/vb/tests/vb/test_vb_arrays_advanced_ops.rs

Module M
    Sub Main()
        Dim src() As Integer = {1, 2, 3, 4, 5}
        Dim dest(4) As Integer
        
        System.Array.Copy(src, 1, dest, 2, 2)
        
        For Each v In dest
            Console.WriteLine(v)
        Next
    End Sub
End Module
