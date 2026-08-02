' vybe-test: vb/vb_objects_collections/c27_dictionary_iterate_values
' origin: languages/vb/tests/vb/vb_objects_collections_test.rs

Dim dict As New Dictionary(Of String, Integer)
dict.Add("a", 10)
dict.Add("b", 20)
Dim vals = dict.Values()
Dim total As Integer = 0
For Each v As Integer In vals
    total = total + v
Next
Console.WriteLine(total)
