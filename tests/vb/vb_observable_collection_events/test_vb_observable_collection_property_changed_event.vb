' vybe-test: vb/vb_observable_collection_events/test_vb_observable_collection_property_changed_event
' origin: languages/vb/tests/vb/test_vb_observable_collection_events.rs

Imports System.Collections.ObjectModel
Imports System.ComponentModel

Module Program
    Sub Main()
        Dim collection As New ObservableCollection(Of Integer)()
        AddHandler CType(collection, INotifyPropertyChanged).PropertyChanged, Sub(sender, e)
            Console.WriteLine("Prop: " & e.PropertyName)
        End Sub

        collection.Add(100)
    End Sub
End Module
