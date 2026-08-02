' vybe-test: vb/vb_for_each_custom_enumerator_struct/test_vb_for_each_dictionary_key_value_pair
' origin: languages/vb/tests/vb/test_vb_for_each_custom_enumerator_struct.rs

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer)()
        dict("A") = 1
        dict("B") = 2

        For Each kvp As KeyValuePair(Of String, Integer) In dict
            Console.WriteLine(kvp.Key & ":" & kvp.Value)
        Next
    End Sub
End Module
