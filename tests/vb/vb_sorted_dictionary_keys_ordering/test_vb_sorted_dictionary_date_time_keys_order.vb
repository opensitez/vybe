' vybe-test: vb/vb_sorted_dictionary_keys_ordering/test_vb_sorted_dictionary_date_time_keys_order
' origin: languages/vb/tests/vb/test_vb_sorted_dictionary_keys_ordering.rs

Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of DateTime, String)()
        dict(New DateTime(2025, 12, 31)) = "New Year Eve"
        dict(New DateTime(2025, 1, 1)) = "New Year Day"
        dict(New DateTime(2025, 6, 15)) = "Mid Year"

        Dim dates As New List(Of String)()
        For Each d In dict.Keys
            dates.Add(d.ToString("yyyy-MM-dd"))
        Next
        Console.WriteLine(String.Join(",", dates))
    End Sub
End Module
