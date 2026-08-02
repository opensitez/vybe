' vybe-test: vb/vb_dictionary_key_collection/test_vb_dict_kvp_deconstruction
' origin: languages/vb/tests/vb/test_vb_dictionary_key_collection.rs

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer) From {{"Key1", 100}}
        For Each kvp In dict
            Console.WriteLine(kvp.Key & ":" & kvp.Value)
        Next
    End Sub
End Module
