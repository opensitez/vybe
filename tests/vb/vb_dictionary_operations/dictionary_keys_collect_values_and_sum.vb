' vybe-test: vb/vb_dictionary_operations/dictionary_keys_collect_values_and_sum
' origin: languages/vb/tests/vb/test_vb_dictionary_operations.rs

Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map.Add("x", 10)
        map.Add("y", 20)
        map.Add("z", 30)

        Dim total As Integer = 0
        For Each key As String In map.Keys
            total += map(key)
        Next

        Console.WriteLine(map.Keys.Count)
        Console.WriteLine(total)
    End Sub
End Module
