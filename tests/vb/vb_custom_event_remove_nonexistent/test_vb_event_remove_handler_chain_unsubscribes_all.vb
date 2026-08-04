' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_remove_handler_chain_unsubscribes_all
' origin: languages/vb/tests/vb/test_vb_custom_event_remove_nonexistent.rs

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

Imports System

Class Emitter
    Public Event Data As Action(Of Integer)
    Public Sub Push(v As Integer)
        RaiseEvent Data(v)
    End Sub
End Class

Module Program
    Sub Main()
        Dim e As New Emitter()
        Dim h1 As Action(Of Integer) = Sub(v) __P(CStr("H1:" & v))
        Dim h2 As Action(Of Integer) = Sub(v) __P(CStr("H2:" & v))

        AddHandler e.Data, h1
        AddHandler e.Data, h2
        e.Push(1)
        RemoveHandler e.Data, h1
        RemoveHandler e.Data, h2
        e.Push(2)
        __P(CStr("Done"))
        __Check("H1:1
H2:1
Done")
    End Sub
End Module
