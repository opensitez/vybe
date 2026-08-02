' vybe-test: vb/vb_tuple_deconstruct_method_overloads/test_vb_tuple_array_declaration_and_iteration
' origin: languages/vb/tests/vb/test_vb_tuple_deconstruct_method_overloads.rs

Module Program
    Sub Main()
        Dim pairs As (Key As String, Val As Integer)() = {("A", 1), ("B", 2)}
        For Each p In pairs
            Console.WriteLine(p.Key & "=" & p.Val)
        Next
    End Sub
End Module
