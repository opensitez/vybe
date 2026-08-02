' vybe-test: vb/vb_property_changed_event_firing/test_vb_property_changed_reentrant_property_update
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

Class ReentrantVM
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _val1 As Integer
    Private _val2 As Integer

    Public Property Val1 As Integer
        Get
            Return _val1
        End Get
        Set(v As Integer)
            _val1 = v
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Val1"))
        End Set
    End Property

    Public Property Val2 As Integer
        Get
            Return _val2
        End Get
        Set(v As Integer)
            _val2 = v
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Val2"))
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim vm As New ReentrantVM()
        AddHandler vm.PropertyChanged, Sub(s, e)
            If e.PropertyName = "Val1" Then
                vm.Val2 = vm.Val1 * 10
            End If
        End Sub

        vm.Val1 = 5
        __Check(CStr(vm.Val2), "50")
    End Sub
End Module
