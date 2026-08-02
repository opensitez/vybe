' vybe-test: vb/vb_collections_advanced/list_foreach
' origin: languages/vb/tests/vb/test_vb_collections_advanced.rs

Module M
    Sub Main()
        Dim items As New List(Of String)
        items.Add("a")
        items.Add("b")
        items.Add("c")
        Dim result As String = ""
        For Each item As String In items
            result = result & item
        Next
        Console.WriteLine(result)
    End Sub
End Module
