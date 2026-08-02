' vybe-test: vb/vb_multidimensional_array_slicing/test_vb_array_2d_transpose
' origin: languages/vb/tests/vb/test_vb_multidimensional_array_slicing.rs

Module Program
    Sub Main()
        Dim orig(,) As Integer = {{1, 2, 3}, {4, 5, 6}}
        Dim trans(2, 1) As Integer
        For r As Integer = 0 To 1
            For c As Integer = 0 To 2
                trans(c, r) = orig(r, c)
            Next
        Next
        Console.WriteLine(trans(0, 1))
        Console.WriteLine(trans(2, 0))
    End Sub
End Module
