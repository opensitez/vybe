' vybe-test: vb/vb_property_changed_event_firing/test_vb_property_changed_nullable_type_property
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

Class NullableViewModel
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _date As DateTime?
    Public Property ExpiryDate As DateTime?
        Get
            Return _date
        End Get
        Set(value As DateTime?)
            _date = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("ExpiryDate"))
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim vm As New NullableViewModel()
        AddHandler vm.PropertyChanged, Sub(s, e) __Check(CStr("Date Changed"), "Date Changed")
        vm.ExpiryDate = New DateTime(2030, 1, 1)
    End Sub
End Module
