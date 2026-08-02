' vybe-test: vb/vb_key_value_pair_struct_usage/test_vb_key_value_pair_collection_projection
' origin: languages/vb/tests/vb/test_vb_key_value_pair_struct_usage.rs

Imports System.Collections.Generic
Imports System.Linq

Module Program
    Sub Main()
        Dim keys = {"K1", "K2", "K3"}
        Dim vals = {10, 20, 30}
        Dim pairs = keys.Zip(vals, Function(k, v) New KeyValuePair(Of String, Integer)(k, v))
        For Each p In pairs
            Console.WriteLine(p.Key & "=" & p.Value)
        Next
    End Sub
End Module
