' vybe-test: vb/vb_for_each_custom_enumerator_struct/test_vb_for_each_array_implicit_type_inference
' origin: languages/vb/tests/vb/test_vb_for_each_custom_enumerator_struct.rs

Module Program
    Sub Main()
        Dim items As String() = {"Alpha", "Beta", "Gamma"}
        Dim concat = ""
        For Each item In items
            concat &= item
        Next
        Console.WriteLine(concat)
    End Sub
End Module
