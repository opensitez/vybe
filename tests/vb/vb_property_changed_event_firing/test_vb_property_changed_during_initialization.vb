' vybe-test: vb/vb_property_changed_event_firing/test_vb_property_changed_during_initialization
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

Class InitializingVM
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Public Property Title As String

    Public Sub New(t As String)
        ' Subscribing inside constructor or firing after constructor
        Title = t
    End Sub

    Public Sub InitDone()
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Title"))
    End Sub
End Class

Module Program
    Sub Main()
        Dim vm As New InitializingVM("InitTitle")
        AddHandler vm.PropertyChanged, Sub(s, e) __Check(CStr(e.PropertyName & "=" & vm.Title), "Title=InitTitle")
        vm.InitDone()
    End Sub
End Module
