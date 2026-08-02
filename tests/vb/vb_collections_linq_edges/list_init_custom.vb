' vybe-test: vb/vb_collections_linq_edges/list_init_custom
' origin: languages/vb/tests/vb/test_vb_collections_linq_edges.rs

Imports System.Collections.Generic: Class C: Implements IEnumerable: Public Sub Add(x As Integer): Console.WriteLine(x): End Sub: Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator: Return Nothing: End Function: End Class: Module M: Sub Main(): Dim c As New C From {10, 20}: End Sub: End Module
