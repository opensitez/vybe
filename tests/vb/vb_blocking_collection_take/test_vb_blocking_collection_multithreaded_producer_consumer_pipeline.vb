' vybe-test: vb/vb_blocking_collection_take/test_vb_blocking_collection_multithreaded_producer_consumer_pipeline
' origin: languages/vb/tests/vb/test_vb_blocking_collection_take.rs

Imports System.Collections.Concurrent
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim bc As New BlockingCollection(Of Integer)()

        Dim producer = Task.Run(Sub()
            For i As Integer = 1 To 5 : bc.Add(i) : Next
            bc.CompleteAdding()
        End Sub)

        Dim consumerSum = 0
        Dim consumer = Task.Run(Sub()
            For Each item In bc.GetConsumingEnumerable()
                consumerSum += item
            Next
        End Sub)

        Task.WaitAll(producer, consumer)
        Console.WriteLine("Consumer Sum: " & consumerSum)
    End Sub
End Module
