' vybe-test: vb/vb_objects_collections/b12_list_for_each_order
' origin: languages/vb/tests/vb/vb_objects_collections_test.rs

Dim list As New List(Of String)
list.Add("a")
list.Add("b")
list.Add("c")
For Each item As String In list
    Console.WriteLine(item)
Next
