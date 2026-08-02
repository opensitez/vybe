' vybe-test: vb/vb_named_tuple_nested_structures/test_vb_named_tuple_linq_select_projection
' origin: languages/vb/tests/vb/test_vb_named_tuple_nested_structures.rs

Imports System.Linq

Module Program
    Sub Main()
        Dim nums = {1, 2, 3}
        Dim tuples = From n In nums Select Entry = (Value:=n, Square:=n * n)
        For Each t In tuples
            Console.WriteLine(t.Value & "^2=" & t.Square)
        Next
    End Sub
End Module
