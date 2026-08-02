' vybe-test: vb/vb_blocking_collection_producer_consumer/test_vb_blocking_collection_complete_adding_consuming_enumerable
' origin: languages/vb/tests/vb/test_vb_blocking_collection_producer_consumer.rs

Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bc As New BlockingCollection(Of Integer)()
        bc.Add(10)
        bc.Add(20)
        bc.CompleteAdding()

        Console.WriteLine(bc.IsAddingCompleted)
        Dim sum As Integer = 0
        For Each val In bc.GetConsumingEnumerable()
            sum += val
        Next
        Console.WriteLine(sum)
        Console.WriteLine(bc.IsCompleted)
    End Sub
End Module
