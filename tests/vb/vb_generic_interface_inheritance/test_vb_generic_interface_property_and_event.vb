' vybe-test: vb/vb_generic_interface_inheritance/test_vb_generic_interface_property_and_event
' origin: languages/vb/tests/vb/test_vb_generic_interface_inheritance.rs

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

Imports System

Interface IObservableValue(Of T)
    Property Value As T
    Event ValueChanged As Action(Of T)
End Interface

Class ObservableInt
    Implements IObservableValue(Of Integer)
    Public Event ValueChanged As Action(Of Integer) Implements IObservableValue(Of Integer).ValueChanged
    Private _val As Integer
    Public Property Value As Integer Implements IObservableValue(Of Integer).Value
        Get
            Return _val
        Get
        End Get
        Set(val As Integer)
            _val = val
            RaiseEvent ValueChanged(_val)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim obs As IObservableValue(Of Integer) = New ObservableInt()
        AddHandler obs.ValueChanged, Sub(v) __Check(CStr("New Value: " & v), "New Value: 42")
        obs.Value = 42
    End Sub
End Module
