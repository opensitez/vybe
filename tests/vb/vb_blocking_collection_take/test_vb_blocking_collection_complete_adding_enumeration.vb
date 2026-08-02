' vybe-test: vb/vb_blocking_collection_take/test_vb_blocking_collection_complete_adding_enumeration
' origin: languages/vb/tests/vb/test_vb_blocking_collection_take.rs

Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bc As New BlockingCollection(Of Integer)()
        bc.Add(10)
        bc.Add(20)
        bc.CompleteAdding()

        Dim list As New System.Collections.Generic.List(Of Integer)()
        For Each item In bc.GetConsumingEnumerable()
            list.Add(item)
        Next
        Console.WriteLine(String.Join(",", list) & "|IsCompleted=" & bc.IsCompleted)
    End Sub
End Module
