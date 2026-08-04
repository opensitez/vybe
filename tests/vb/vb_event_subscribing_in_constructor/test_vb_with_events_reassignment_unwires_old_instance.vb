' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_with_events_reassignment_unwires_old_instance
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
    Public Property Name As String
    Public Event Action As EventHandler
    Public Sub Fire()
        RaiseEvent Action(Me, EventArgs.Empty)
    End Sub
End Class

Class SwitchableListener
    Private WithEvents currentEmitter As Emitter

    Public Sub SetEmitter(e As Emitter)
        currentEmitter = e ' Unwires previous currentEmitter, wires new e!
    End Sub

    Private Sub OnAction(sender As Object, e As EventArgs) Handles currentEmitter.Action
        __P(CStr("Action Handled From: " & currentEmitter.Name))
    End Sub
End Class

Module Program
    Sub Main()
        Dim e1 As New Emitter With {.Name = "First"}
        Dim e2 As New Emitter With {.Name = "Second"}

        Dim listener As New SwitchableListener()
        listener.SetEmitter(e1)
        e1.Fire()

        listener.SetEmitter(e2)
        e1.Fire() ' Should NOT fire listener!
        e2.Fire() ' Should fire listener!
        __Check("Action Handled From: First
Action Handled From: Second")
    End Sub
End Module
