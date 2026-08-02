' vybe-test: vb/vb_list_collection_matrix/list_foreach_sums_values
' origin: languages/vb/tests/vb/test_vb_list_collection_matrix.rs

Imports System.Collections.Generic

Module M
    Sub Main()
        Dim values As New List(Of Integer) From {1, 2, 3}
        Dim total As Integer = 0
        For Each value As Integer In values
            total += value
        Next
        Console.WriteLine(total)
    End Sub
End Module
