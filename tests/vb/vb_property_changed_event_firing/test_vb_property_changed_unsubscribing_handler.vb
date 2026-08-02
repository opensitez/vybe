' vybe-test: vb/vb_property_changed_event_firing/test_vb_property_changed_unsubscribing_handler
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

Class Target
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _val As Integer
    Public Property Value As Integer
        Get
            Return _val
        End Get
        Set(v As Integer)
            _val = v
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Value"))
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim t As New Target()
        Dim count = 0
        Dim handler As PropertyChangedEventHandler = Sub(s, e) count += 1

        AddHandler t.PropertyChanged, handler
        t.Value = 1
        RemoveHandler t.PropertyChanged, handler
        t.Value = 2
        __Check(CStr(count), "1")
    End Sub
End Module
