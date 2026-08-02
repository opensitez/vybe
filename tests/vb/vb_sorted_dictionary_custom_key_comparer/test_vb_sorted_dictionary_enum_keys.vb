' vybe-test: vb/vb_sorted_dictionary_custom_key_comparer/test_vb_sorted_dictionary_enum_keys
' origin: languages/vb/tests/vb/test_vb_sorted_dictionary_custom_key_comparer.rs

Imports System.Collections.Generic

Enum Priority
    Low = 0
    Medium = 1
    High = 2
End Enum

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Priority, String)()
        dict(Priority.High) = "Emergency"
        dict(Priority.Low) = "Routine"
        For Each kv In dict
            Console.WriteLine(kv.Key.ToString() & ":" & kv.Value)
        Next
    End Sub
End Module
