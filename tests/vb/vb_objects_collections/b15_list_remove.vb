' vybe-test: vb/vb_objects_collections/b15_list_remove
' origin: languages/vb/tests/vb/vb_objects_collections_test.rs

Dim list As New List(Of String)
list.Add("a")
list.Add("b")
list.Add("c")
list.Remove("b")
Console.WriteLine(list.Count)
For Each item As String In list
    Console.WriteLine(item)
Next
