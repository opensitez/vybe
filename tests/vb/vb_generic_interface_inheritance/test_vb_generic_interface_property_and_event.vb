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
'
' Output is COLLECTED, not paired. The emitter rewrites every
' `Console.WriteLine(x)` into `__P(CStr(x))` and compares the whole output once
' at the end of `Sub Main`. Pairing the i-th print with the i-th expected line
' cannot assert anything about a loop, and loops alone were 402 of VB's 6,671
' cases.
'
' Rendering happens at the CALL SITE via `CStr`, where the expression still has
' its static type — the same reason the C# harness renders with `.ToString()`
' rather than inside the helper.

Imports System
Module VybeCheck
    Public __buf As String = ""

    Sub __P(s As String)
        __buf = __buf & s & vbLf
    End Sub

    Sub __Pr(s As String)
        __buf = __buf & s
    End Sub

    ' The final WriteLine contributes a trailing newline that the expected line
    ' vector never carried, so BOTH forms are accepted.
    Sub __Check(want As String)
        If __buf <> want AndAlso __buf <> want & vbLf Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & __buf & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module


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
        AddHandler obs.ValueChanged, Sub(v) __P(CStr("New Value: " & v))
        obs.Value = 42
        __Check("New Value: 42")
    End Sub
End Module
