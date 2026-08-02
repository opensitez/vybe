' vybe-test: vb/vb_tuple_hashset_dictionary_keys/test_vb_tuple_dictionary_iteration
' origin: languages/vb/tests/vb/test_vb_tuple_hashset_dictionary_keys.rs

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of (Integer, Integer), Integer)()
        dict((0, 0)) = 100
        dict((1, 1)) = 200

        For Each kv In dict
            Console.WriteLine(kv.Key.Item1 & "," & kv.Key.Item2 & "=" & kv.Value)
        Next
    End Sub
End Module
