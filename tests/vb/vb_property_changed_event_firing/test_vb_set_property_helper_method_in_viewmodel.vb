' vybe-test: vb/vb_property_changed_event_firing/test_vb_set_property_helper_method_in_viewmodel
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
Imports System.Runtime.CompilerServices

Class BindableBase
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Protected Function SetProperty(Of T)(ByRef storage As T, value As T, <CallerMemberName> Optional propName As String = Nothing) As Boolean
        If EqualityComparer(Of T).Default.Equals(storage, value) Then Return False
        storage = value
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propName))
        Return True
    End Function
End Class

Class CustomerViewModel
    Inherits BindableBase

    Private _age As Integer
    Public Property Age As Integer
        Get
            Return _age
        End Get
        Set(value As Integer)
            SetProperty(_age, value)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim vm As New CustomerViewModel()
        Dim changedName = ""
        AddHandler vm.PropertyChanged, Sub(s, e) changedName = e.PropertyName
        Dim res = vm.Age = 30
        __Check(CStr(changedName & "|Value=" & vm.Age), "Age|Value=30")
    End Sub
End Module
