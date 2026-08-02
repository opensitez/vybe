' vybe-test: vb/vb_property_changed_event_firing/test_vb_property_changed_collection_property_reference_change
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

Imports System.Collections.Generic
Imports System.ComponentModel

Class ListViewModel
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _items As List(Of String)
    Public Property Items As List(Of String)
        Get
            Return _items
        End Get
        Set(value As List(Of String))
            _items = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Items"))
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim vm As New ListViewModel()
        AddHandler vm.PropertyChanged, Sub(s, e) __Check(CStr("List Replaced"), "List Replaced")
        vm.Items = New List(Of String) From {"A", "B"}
    End Sub
End Module
