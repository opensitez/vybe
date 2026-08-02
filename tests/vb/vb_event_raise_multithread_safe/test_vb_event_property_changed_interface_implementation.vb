' vybe-test: vb/vb_event_raise_multithread_safe/test_vb_event_property_changed_interface_implementation
' origin: languages/vb/tests/vb/test_vb_event_raise_multithread_safe.rs

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

Class Person
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Private _name As String
    Public Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            If _name <> value Then
                _name = value
                RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Name"))
            End If
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim p As New Person()
        AddHandler p.PropertyChanged, Sub(s, e) __Check(CStr("Changed: " & e.PropertyName), "Changed: Name")
        p.Name = "Bob"
    End Sub
End Module
