' vybe-test: vb/vb_observable_collection_reset/test_vb_observable_collection_item_property_changed_bubbling
' origin: languages/vb/tests/vb/test_vb_observable_collection_reset.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Imports System.Collections.ObjectModel
Imports System.ComponentModel

Class NotifyingItem
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _val As String
    Public Property Val As String
        Get
            Return _val
        End Get
        Set(v As String)
            _val = v
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Val"))
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim item As New NotifyingItem()
        Dim col As New ObservableCollection(Of NotifyingItem)()
        col.Add(item)

        Dim changed = False
        AddHandler item.PropertyChanged, Sub(s, e) changed = True
        item.Val = "Updated"
        __Check(CStr(changed), "True")
    End Sub
End Module
