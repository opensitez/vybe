' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_virtual_method_event_handler_subscription_in_constructor
' origin: languages/vb/tests/vb/test_vb_event_subscribing_in_constructor.rs

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


Class BaseSubscriber
    Public Sub New(pub As Source)
        AddHandler pub.Ping, AddressOf OnPing
    End Sub

    Protected Overridable Sub OnPing(sender As Object, e As EventArgs)
        __P(CStr("Base OnPing"))
    End Sub
End Class

Class DerivedSubscriber
    Inherits BaseSubscriber

    Public Sub New(pub As Source)
        MyBase.New(pub)
    End Sub

    Protected Overrides Sub OnPing(sender As Object, e As EventArgs)
        __P(CStr("Derived OnPing"))
    End Sub
End Class

Class Source
    Public Event Ping As EventHandler
    Public Sub Fire()
        RaiseEvent Ping(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As New Source()
        Dim subObj As New DerivedSubscriber(s)
        s.Fire()
        __Check("Derived OnPing")
    End Sub
End Module
