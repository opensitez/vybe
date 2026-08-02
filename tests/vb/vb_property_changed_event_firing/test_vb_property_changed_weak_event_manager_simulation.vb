' vybe-test: vb/vb_property_changed_event_firing/test_vb_property_changed_weak_event_manager_simulation
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

Class TargetVM
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Public Sub Notify()
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Data"))
    End Sub
End Class

Module Program
    Sub Main()
        Dim vm As New TargetVM()
        AddHandler vm.PropertyChanged, Sub(s, e) __Check(CStr("Weak Handled: " & e.PropertyName), "Weak Handled: Data")
        vm.Notify()
    End Sub
End Module
