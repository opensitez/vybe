' vybe-test: vb/vb_system_dictionary_matrix/sorted_dictionary_ordering_contract
' origin: languages/vb/tests/vb/test_vb_system_dictionary_matrix.rs

Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New SortedDictionary(Of Integer, String)()
        map.Add(3, "three")
        map.Add(1, "one")
        map.Add(2, "two")
        Dim first As Integer = 0
        Dim second As Integer = 0
        Dim third As Integer = 0
        Dim i As Integer = 0
        For Each value As Integer In map.Keys
            If i = 0 Then
                first = value
            ElseIf i = 1 Then
                second = value
            ElseIf i = 2 Then
                third = value
            End If
            i += 1
        Next
        Console.WriteLine(first)
        Console.WriteLine(second)
        Console.WriteLine(third)
    End Sub
End Module
