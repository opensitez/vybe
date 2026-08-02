' vybe-test: vb/vb_system_collection_api_matrix/collection_api_matrix_dictionary_enumeration_order_stable
' origin: languages/vb/tests/vb/test_vb_system_collection_api_matrix.rs

Imports System
Imports System.Collections.Generic
Imports System.Text

Module M
    Sub Main()
        Dim map As New Dictionary(Of Integer, String)()
        map.Add(2, "b")
        map.Add(1, "a")
        map.Add(3, "c")

        Dim sb As New StringBuilder()
        For Each pair In map
            sb.Append(pair.Key).Append(":").Append(pair.Value).Append(",")
        Next

        Console.WriteLine(sb.ToString().Contains("1:a"))
        Console.WriteLine(sb.ToString().Contains("3:c"))
    End Sub
End Module
