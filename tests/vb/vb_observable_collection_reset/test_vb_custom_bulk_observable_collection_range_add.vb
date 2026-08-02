' vybe-test: vb/vb_observable_collection_reset/test_vb_custom_bulk_observable_collection_range_add
' origin: languages/vb/tests/vb/test_vb_observable_collection_reset.rs

Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized

Class RangeObservableCollection(Of T)
    Inherits ObservableCollection(Of T)

    Public Sub AddRange(items As IEnumerable(Of T))
        For Each item In items
            Items.Add(item)
        Next
        OnCollectionChanged(New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))
    End Sub
End Class

Module Program
    Sub Main()
        Dim col As New RangeObservableCollection(Of Integer)()
        Dim resetCount = 0
        AddHandler col.CollectionChanged, Sub(s, e)
            If e.Action = NotifyCollectionChangedAction.Reset Then resetCount += 1
        End Sub
        col.AddRange({10, 20, 30})
        Console.WriteLine(col.Count & "|Reset=" & resetCount)
    End Sub
End Module
