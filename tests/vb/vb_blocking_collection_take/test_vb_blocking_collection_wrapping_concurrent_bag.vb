' vybe-test: vb/vb_blocking_collection_take/test_vb_blocking_collection_wrapping_concurrent_bag
' origin: languages/vb/tests/vb/test_vb_blocking_collection_take.rs

Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bag As New ConcurrentBag(Of String)()
        Dim bc As New BlockingCollection(Of String)(bag)
        bc.Add("BagItem1")
        bc.Add("BagItem2")

        Dim count = 0
        While bc.Count > 0
            Dim item = bc.Take()
            count += 1
        End While
        Console.WriteLine("Bag Collection Cleared Count: " & count)
    End Sub
End Module
