' vybe-test: vb/vb_property_changed_event_firing/test_vb_property_changed_thread_safe_raise_event
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

Class SafeVM
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private Sub OnPropertyChanged(name As String)
        Dim handler = PropertyChangedEvent
        If handler IsNot Nothing Then
            handler(Me, New PropertyChangedEventArgs(name))
        End If
    End Sub

    Public Sub Trigger(name As String)
        OnPropertyChanged(name)
    End Sub
End Class

Module Program
    Sub Main()
        Dim vm As New SafeVM()
        AddHandler vm.PropertyChanged, Sub(s, e) __Check(CStr("Safe: " & e.PropertyName), "Safe: SafeProp")
        vm.Trigger("SafeProp")
    End Sub
End Module
