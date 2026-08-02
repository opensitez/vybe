' vybe-test: vb/vb_property_changed_event_firing/test_vb_property_changed_nested_viewmodel_event_forwarding
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

Class ChildVM
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Public Sub Fire()
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("ChildProp"))
    End Sub
End Class

Class ParentVM
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Public Property Child As ChildVM

    Public Sub New()
        Child = New ChildVM()
        AddHandler Child.PropertyChanged, Sub(s, e)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Child." & e.PropertyName))
        End Sub
    End Sub
End Class

Module Program
    Sub Main()
        Dim parent As New ParentVM()
        AddHandler parent.PropertyChanged, Sub(s, e) __Check(CStr(e.PropertyName), "Child.ChildProp")
        parent.Child.Fire()
    End Sub
End Module
