' vybe-test: vb/vb_objects_collections/e37_array_for_each
' origin: languages/vb/tests/vb/vb_objects_collections_test.rs

Dim arr(2) As Integer
arr(0) = 1
arr(1) = 2
arr(2) = 3
For Each n As Integer In arr
    Console.WriteLine(n)
Next
