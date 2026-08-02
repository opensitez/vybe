' vybe-test: vb/vb_collections_advanced/array_initializer
' origin: languages/vb/tests/vb/test_vb_collections_advanced.rs

Module M
    Sub Main()
        Dim arr() As Integer = {5, 10, 15, 20}
        Dim sum As Integer = 0
        For Each n As Integer In arr
            sum = sum + n
        Next
        Console.WriteLine(sum)
    End Sub
End Module
