' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_multiple_withevents_fields_in_same_class
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


Class EmitterA
    Public Event EventA As EventHandler
    Public Sub Fire()
        RaiseEvent EventA(Me, EventArgs.Empty)
    End Sub
End Class

Class EmitterB
    Public Event EventB As EventHandler
    Public Sub Fire()
        RaiseEvent EventB(Me, EventArgs.Empty)
    End Sub
End Class

Class CombinedListener
    Public WithEvents SourceA As EmitterA
    Public WithEvents SourceB As EmitterB

    Private Sub OnA(sender As Object, e As EventArgs) Handles SourceA.EventA
        __P(CStr("A Handled"))
    End Sub

    Private Sub OnB(sender As Object, e As EventArgs) Handles SourceB.EventB
        __P(CStr("B Handled"))
    End Sub
End Class

Module Program
    Sub Main()
        Dim l As New CombinedListener()
        l.SourceA = New EmitterA()
        l.SourceB = New EmitterB()
        l.SourceA.Fire()
        l.SourceB.Fire()
        __Check("A Handled
B Handled")
    End Sub
End Module
