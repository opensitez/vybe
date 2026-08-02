' vybe-test: vb/vb_dictionary_operations/dictionary_duplicate_add_raises_key_error
' origin: languages/vb/tests/vb/test_vb_dictionary_operations.rs

Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map.Add("a", 1)
        Try
            map.Add("a", 2)
            Console.WriteLine("Added")
        Catch ex As ArgumentException
            Console.WriteLine("Duplicate")
        End Try
    End Sub
End Module
