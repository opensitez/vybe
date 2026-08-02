' vybe-test: vb/vb_dictionary_operations/dictionary_foreach_over_key_value_pairs
' origin: languages/vb/tests/vb/test_vb_dictionary_operations.rs

Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map.Add("x", 10)
        map.Add("y", 20)
        For Each pair As KeyValuePair(Of String, Integer) In map
            Console.WriteLine(pair.Key & ":" & pair.Value)
        Next
    End Sub
End Module
