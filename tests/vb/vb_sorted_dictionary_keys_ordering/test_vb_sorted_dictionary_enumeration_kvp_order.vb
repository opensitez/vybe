' vybe-test: vb/vb_sorted_dictionary_keys_ordering/test_vb_sorted_dictionary_enumeration_kvp_order
' origin: languages/vb/tests/vb/test_vb_sorted_dictionary_keys_ordering.rs

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, String)()
        dict(2) = "B"
        dict(1) = "A"

        Dim log = ""
        For Each kvp In dict
            log &= kvp.Key & ":" & kvp.Value & ";"
        Next
        Console.WriteLine(log)
    End Sub
End Module
