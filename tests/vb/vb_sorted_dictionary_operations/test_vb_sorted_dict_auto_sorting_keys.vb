' vybe-test: vb/vb_sorted_dictionary_operations/test_vb_sorted_dict_auto_sorting_keys
' origin: languages/vb/tests/vb/test_vb_sorted_dictionary_operations.rs

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, String)
        dict.Add(30, "Thirty")
        dict.Add(10, "Ten")
        dict.Add(20, "Twenty")
        For Each kvp In dict
            Console.WriteLine(kvp.Key & ":" & kvp.Value)
        Next
    End Sub
End Module
