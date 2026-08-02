' vybe-test: vb/vb_tuple_hashset_dictionary_keys/test_vb_tuple_dictionary_values_collection
' origin: languages/vb/tests/vb/test_vb_tuple_hashset_dictionary_keys.rs

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of Integer, (Name As String, Age As Integer)) From {
            {1, ("Alice", 25)},
            {2, ("Bob", 30)}
        }
        For Each val In dict.Values
            Console.WriteLine(val.Name & "=" & val.Age)
        Next
    End Sub
End Module
