' vybe-test: vb/vb_key_value_pair_struct_usage/test_vb_key_value_pair_array_iteration
' origin: languages/vb/tests/vb/test_vb_key_value_pair_struct_usage.rs

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim pairs As KeyValuePair(Of String, Integer)() = {
            New KeyValuePair(Of String, Integer)("A", 1),
            New KeyValuePair(Of String, Integer)("B", 2)
        }
        For Each pair In pairs
            Console.WriteLine(pair.Key & ":" & pair.Value)
        Next
    End Sub
End Module
