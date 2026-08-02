' vybe-test: vb/vb_objects_collections/b18_list_of_numbers_sum
' origin: languages/vb/tests/vb/vb_objects_collections_test.rs

Dim list As New List(Of Integer)
list.Add(10)
list.Add(20)
list.Add(30)
Dim total As Integer = 0
For Each n As Integer In list
    total = total + n
Next
Console.WriteLine(total)
