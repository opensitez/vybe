' vybe-test: vb/vb_property_changed_event_firing/test_vb_property_changed_indexed_property_notification
' origin: languages/vb/tests/vb/test_vb_property_changed_event_firing.rs

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

Imports System.ComponentModel

Class IndexedModel
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private data(2) As String
    Default Public Property Item(idx As Integer) As String
        Get
            Return data(idx)
        End Get
        Set(value As String)
            data(idx) = value
            ' Signal indexed property change via "Item[]" or "Item"
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Item[]"))
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim m As New IndexedModel()
        AddHandler m.PropertyChanged, Sub(s, e) __Check(CStr(e.PropertyName), "Item[]")
        m(0) = "Val1"
    End Sub
End Module
