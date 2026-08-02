' vybe-test: vb/vb_dictionary_contains_value_lookup/test_vb_dictionary_add_duplicate_key_throws_or_handled
' origin: languages/vb/tests/vb/test_vb_dictionary_contains_value_lookup.rs

Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer)()
        dict.Add("Unique", 1)
        Try
            dict.Add("Unique", 2)
            Console.WriteLine("Added Duplicate")
        Catch ex As ArgumentException
            Console.WriteLine("ArgumentException")
        End Try
    End Sub
End Module
