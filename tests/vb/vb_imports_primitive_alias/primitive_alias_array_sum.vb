' vybe-test: vb/vb_imports_primitive_alias/primitive_alias_array_sum
' origin: languages/vb/tests/vb/test_vb_imports_primitive_alias.rs

Imports MyInt = System.Int32

Module M
    Sub Main()
        Dim values() As MyInt = {1, 2, 3, 4}
        Dim total As MyInt = 0
        For Each item As MyInt In values
            total = total + item
        Next
        Console.WriteLine(total)
    End Sub
End Module
