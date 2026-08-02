' vybe-test: vb/vb_concurrent_dictionary_try_add/test_vb_concurrent_dictionary_enumeration_during_mutation_safe
' origin: languages/vb/tests/vb/test_vb_concurrent_dictionary_try_add.rs

Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of Integer, String)()
        dict(1) = "One"
        dict(2) = "Two"

        Dim count = 0
        For Each kvp In dict
            count += 1
            dict(count + 10) = "Extra" ' Safe to mutate during enumeration!
        Next
        Console.WriteLine(count >= 2)
    End Sub
End Module
