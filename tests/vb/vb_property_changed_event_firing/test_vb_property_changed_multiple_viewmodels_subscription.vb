' vybe-test: vb/vb_property_changed_event_firing/test_vb_property_changed_multiple_viewmodels_subscription
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

Class SimpleModel
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Public Sub Touch(name As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(name))
    End Sub
End Class

Module Program
    Sub Main()
        Dim m1 As New SimpleModel()
        Dim m2 As New SimpleModel()
        AddHandler m1.PropertyChanged, Sub(s, e) __Check(CStr("M1: " & e.PropertyName), "M1: P1")
        AddHandler m2.PropertyChanged, Sub(s, e) __Check(CStr("M2: " & e.PropertyName), "M2: P2")
        m1.Touch("P1")
        m2.Touch("P2")
    End Sub
End Module
