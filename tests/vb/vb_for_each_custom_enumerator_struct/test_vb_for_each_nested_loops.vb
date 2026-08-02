' vybe-test: vb/vb_for_each_custom_enumerator_struct/test_vb_for_each_nested_loops
' origin: languages/vb/tests/vb/test_vb_for_each_custom_enumerator_struct.rs

Module Program
    Sub Main()
        Dim letters As String() = {"X", "Y"}
        Dim numbers As Integer() = {1, 2}
        For Each l In letters
            For Each n In numbers
                Console.WriteLine(l & n)
            Next
        Next
    End Sub
End Module
