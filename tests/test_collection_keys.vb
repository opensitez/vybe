Imports System
Module TestCollectionKeys
    Dim passed As Integer = 0
    Dim failed As Integer = 0

    Sub Assert(condition As Boolean, msg As String)
        If condition Then
            passed += 1
        Else
            Console.WriteLine("FAIL: " & msg)
            failed += 1
        End If
    End Sub

    Sub Main()
        Console.WriteLine("=== Collection Key-Based Access Tests ===")
        Console.WriteLine()

        ' ----- 1. Basic Add with key and retrieve by key -----
        Dim col As New Collection()
        col.Add("Alice", "first")
        col.Add("Bob", "second")
        col.Add("Charlie", "third")
        Assert(col.Item("first") = "Alice", "Item by key 'first'")
        Assert(col.Item("second") = "Bob", "Item by key 'second'")
        Assert(col.Item("third") = "Charlie", "Item by key 'third'")

        ' ----- 2. Case-insensitive key lookup -----
        Assert(col.Item("FIRST") = "Alice", "Key lookup is case-insensitive")
        Assert(col.Item("Second") = "Bob", "Key lookup mixed case")

        ' ----- 3. Integer index still works -----
        Assert(col.Item(0) = "Alice", "Item by index 0")
        Assert(col.Item(1) = "Bob", "Item by index 1")
        Assert(col.Item(2) = "Charlie", "Item by index 2")

        ' ----- 4. String indexer via ArrayAccess: col("key") -----
        Assert(col("first") = "Alice", "col(""first"") string indexer")
        Assert(col("third") = "Charlie", "col(""third"") string indexer")

        ' ----- 5. Integer indexer via ArrayAccess -----
        Assert(col(0) = "Alice", "col(0) integer indexer")
        Assert(col(1) = "Bob", "col(1) integer indexer")

        ' ----- 6. Count is correct -----
        Assert(col.Count = 3, "Count = 3")

        ' ----- 7. ContainsKey -----
        Assert(col.ContainsKey("first") = True, "ContainsKey first -> True")
        Assert(col.ContainsKey("SECOND") = True, "ContainsKey SECOND -> True")
        Assert(col.ContainsKey("missing") = False, "ContainsKey missing -> False")

        ' ----- 8. Remove by key -----
        col.Remove("second")
        Assert(col.Count = 2, "After Remove by key, Count = 2")
        Assert(col.ContainsKey("second") = False, "Key 'second' removed")
        Assert(col.Item("first") = "Alice", "first still accessible after remove")
        Assert(col.Item("third") = "Charlie", "third still accessible after remove")
        Assert(col(0) = "Alice", "Index 0 = Alice after remove")
        Assert(col(1) = "Charlie", "Index 1 = Charlie after remove")

        ' ----- 9. Duplicate key detection -----
        Dim dupCaught As Boolean = False
        Try
            Dim colDup As New Collection()
            colDup.Add("A", "key1")
            colDup.Add("B", "key1")  ' Duplicate — should throw
        Catch ex As Exception
            dupCaught = True
        End Try
        Assert(dupCaught = True, "Duplicate key throws exception")

        ' ----- 10. Clear resets keys -----
        Dim col2 As New Collection()
        col2.Add("X", "keyX")
        col2.Add("Y", "keyY")
        col2.Clear()
        Assert(col2.Count = 0, "Clear resets count to 0")
        Assert(col2.ContainsKey("keyX") = False, "Clear removes keys")

        ' ----- 11. Mixed keyed and unkeyed items -----
        Dim col3 As New Collection()
        col3.Add("NoKey1")
        col3.Add("WithKey", "myKey")
        col3.Add("NoKey2")
        Assert(col3.Count = 3, "Mixed keyed/unkeyed count")
        Assert(col3(0) = "NoKey1", "Unkeyed item by index")
        Assert(col3.Item("myKey") = "WithKey", "Keyed item by key")
        Assert(col3(2) = "NoKey2", "Third item by index")

        ' ----- 12. RemoveAt with keys -----
        Dim col4 As New Collection()
        col4.Add("A", "keyA")
        col4.Add("B", "keyB")
        col4.Add("C", "keyC")
        col4.RemoveAt(1)  ' Remove "B"
        Assert(col4.Count = 2, "RemoveAt count")
        Assert(col4.ContainsKey("keyB") = False, "keyB removed by RemoveAt")
        Assert(col4.ContainsKey("keyA") = True, "keyA still valid")
        Assert(col4.ContainsKey("keyC") = True, "keyC still valid")
        Assert(col4.Item("keyA") = "A", "keyA -> A after RemoveAt")
        Assert(col4.Item("keyC") = "C", "keyC -> C after RemoveAt")

        ' ----- 13. For Each on keyed collection -----
        Dim col5 As New Collection()
        col5.Add("Apple", "a")
        col5.Add("Banana", "b")
        col5.Add("Cherry", "c")
        Dim result As String = ""
        For Each item As String In col5
            result &= item & ","
        Next
        Assert(result = "Apple,Banana,Cherry,", "For Each iteration order preserved")

        ' ----- Summary -----
        Console.WriteLine()
        Console.WriteLine("Passed: " & passed)
        Console.WriteLine("Failed: " & failed)
        Console.WriteLine("Total:  " & (passed + failed))
    End Sub
End Module
