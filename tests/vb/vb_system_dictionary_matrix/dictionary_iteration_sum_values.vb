' vybe-test: vb/vb_system_dictionary_matrix/dictionary_iteration_sum_values
' origin: languages/vb/tests/vb/test_vb_system_dictionary_matrix.rs

Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map("a") = 10
        map("b") = 20
        map("c") = 30
        Dim total As Integer = 0
        For Each kv As KeyValuePair(Of String, Integer) In map
            total += kv.Value
        Next
        Console.WriteLine(total)
        Console.WriteLine(map.Count)
    End Module
End Module
