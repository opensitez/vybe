' vybe-test: vb/vb_list_collection_matrix/list_reverse_inverts_order
' origin: languages/vb/tests/vb/test_vb_list_collection_matrix.rs

Imports System.Collections.Generic

Module M
    Sub Main()
        Dim values As New List(Of Integer) From {1, 2, 3}
        values.Reverse()
        For Each value As Integer In values
            Console.WriteLine(value)
        Next
    End Sub
End Module
