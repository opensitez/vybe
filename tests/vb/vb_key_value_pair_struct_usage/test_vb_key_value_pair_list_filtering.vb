' vybe-test: vb/vb_key_value_pair_struct_usage/test_vb_key_value_pair_list_filtering
' origin: languages/vb/tests/vb/test_vb_key_value_pair_struct_usage.rs

Imports System.Collections.Generic
Imports System.Linq

Module Program
    Sub Main()
        Dim list As New List(Of KeyValuePair(Of String, Integer)) From {
            New KeyValuePair(Of String, Integer)("High", 90),
            New KeyValuePair(Of String, Integer)("Low", 10),
            New KeyValuePair(Of String, Integer)("High", 85)
        }
        Dim highs = list.Where(Function(kv) kv.Key = "High")
        For Each kv In highs
            Console.WriteLine(kv.Value)
        Next
    End Sub
End Module
