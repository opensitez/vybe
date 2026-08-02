' vybe-test: vb/vb_sorted_dictionary_custom_key_comparer/test_vb_sorted_dictionary_datetime_keys
' origin: languages/vb/tests/vb/test_vb_sorted_dictionary_custom_key_comparer.rs

Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of DateTime, String)()
        dict(New DateTime(2025, 12, 31)) = "New Year's Eve"
        dict(New DateTime(2025, 1, 1)) = "New Year's Day"
        For Each kv In dict
            Console.WriteLine(kv.Key.ToString("yyyy-MM-dd") & "=" & kv.Value)
        Next
    End Sub
End Module
